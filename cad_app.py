import time
import random
import string
import win32com.client
import pythoncom
import numpy as np
import matplotlib.pyplot as plt

from typing import List, Tuple, Set
from segment import Segment
from shapely.geometry import LineString, Polygon
from shapely.ops import polygonize
from scipy.optimize import bisect


class CadApp:
    def load(self):
        for i in range(10):
            try:
                # 尝试获取现有的 AutoCAD 实例
                acad = win32com.client.Dispatch("AutoCAD.Application")
                self._doc = acad.ActiveDocument
                self._msp = self._doc.ModelSpace
            except Exception as e:
                print(f"【提示】无法连接到 AutoCAD，重新连接中: {e}")
                time.sleep(1)
                continue
        if not self._doc:
            return False
        return True


    # 获取 AutoCAD 中当前选中的直线和多段线（使用 pywin32）
    def _get_selected_lines_and_polylines(self, prompt="请选择直线和多段线"):
        try:
            # 获取选择集
            # 生成随机字符串
            random_string = ''.join(random.choices(string.ascii_letters, k=10))
            selection_set = self._doc.SelectionSets.Add(random_string)  # 创建一个临时选择集
        except Exception as e:
            print(f"创建选择集失败: {e}")
            return [], []

        try:
            self._doc.Utility.prompt(prompt)  # 显示提示信息
            selection_set.SelectOnScreen()  # 添加正确的参数格式来处理AutoCAD选择
        except Exception as e:
            print(f"屏幕选择失败: {e}")
            # selection_set.Delete() #删除创建的选择集
            return [], []

        lines = []
        polylines = []
        for i in range(5):  # 最多尝试5次
            try:
                selection_set.Count
                break  # 如果成功获取选择集，跳出循环
            except Exception as e:
                # print(f"获取选择集失败，第{i+1}次尝试: {e}")
                time.sleep(1)  # 等待1秒后重试

        for i in range(selection_set.Count):
            obj = selection_set.Item(i)  # 获取选择集中的每个对象

            if obj.ObjectName == "AcDbLine":
                lines.append(obj)
            elif obj.ObjectName == "AcDbPolyline" :
                polylines.append(obj)
            elif obj.ObjectName == "AcDb2dPolyline":
                polylines.append(obj)
            else:
                print(f"未知对象类型: {obj.ObjectName}")

        selection_set.Delete()  # 删除临时选择集
        return lines, polylines


    def get_boundary_coords(self, prompt):
        lines, polylines = self._get_selected_lines_and_polylines(prompt)
        coords = []
        # 遍历直线，将坐标存入coords
        for line in lines:
            start_point = line.StartPoint[:2]  # 取前两个坐标（x, y）
            end_point = line.EndPoint[:2]
            coords.append([[start_point, end_point]])
        # 遍历多段线，将坐标存入coords
        for polyline in polylines:
            poly_coords = list(polyline.Coordinates)
            if polyline.objectName == "AcDb2dPolyline":
                points = [(poly_coords[i], poly_coords[i + 1]) for i in range(0, len(poly_coords), 3)]
            else:
                points = [(poly_coords[i], poly_coords[i + 1]) for i in range(0, len(poly_coords), 2)]
            #检查是否为闭合多段线
            if polyline.Closed:
                points.append(points[0])
            #point组成线段
            lines = []
            for i in range(len(points) - 1):
                lines.append([points[i], points[i + 1]])
            coords.append(lines)
        return coords


    def _extend_point(self,point, direction, extend_length):
        """
        延长给定点沿着指定方向的长度
        参数:
        point: 原始点坐标 (x, y)
        direction: 方向向量
        extend_length: 延长的长度

        返回:
        新的点坐标
        """
        unit_direction = direction / np.linalg.norm(direction)  # 单位化方向向量
        new_point = np.array(point) + unit_direction * extend_length
        return (float(new_point[0]), float(new_point[1]))

    def is_collinear(self,p1, p2, p3):
        """判断三个点是否共线"""
        x1, y1 = p1
        x2, y2 = p2
        x3, y3 = p3
        # 使用斜率判断，考虑误差
        if abs(x2 - x1) < 1e-10:  # 垂直线的情况
            return abs(x3 - x1) < 1e-10
        if abs(x3 - x2) < 1e-10:  # 垂直线的情况
            return abs(x2 - x1) < 1e-10
        
        slope1 = (y2 - y1) / (x2 - x1)
        slope2 = (y3 - y2) / (x3 - x2)
        return abs(slope1 - slope2) < 1e-10

    def _segments_overlap(self,seg1, seg2):
        """判断两条线段是否重叠（共线且有重叠部分）"""
        (x1, y1), (x2, y2) = seg1
        (x3, y3), (x4, y4) = seg2
        
        # 首先判断是否共线
        if not (self.is_collinear((x1, y1), (x2, y2), (x3, y3)) and 
                self.is_collinear((x1, y1), (x2, y2), (x4, y4))):
            return False
        
        # 如果是垂直线
        if abs(x2 - x1) < 1e-10:
            y_min1, y_max1 = min(y1, y2), max(y1, y2)
            y_min2, y_max2 = min(y3, y4), max(y3, y4)
            return not (y_max1 < y_min2 or y_max2 < y_min1)
        
        # 如果是水平线或斜线
        x_min1, x_max1 = min(x1, x2), max(x1, x2)
        x_min2, x_max2 = min(x3, x4), max(x3, x4)
        return not (x_max1 < x_min2 or x_max2 < x_min1)

    def _merge_two_segments(self,seg1, seg2):
        """合并两条重叠的线段"""
        (x1, y1), (x2, y2) = seg1
        (x3, y3), (x4, y4) = seg2
        
        # 如果是垂直线
        if abs(x2 - x1) < 1e-10:
            points = [(x1, y1), (x2, y2), (x3, y3), (x4, y4)]
            points.sort(key=lambda p: p[1])  # 按y坐标排序
            return [points[0], points[-1]]  # 返回y坐标最小和最大的点
        
        # 如果是水平线或斜线
        points = [(x1, y1), (x2, y2), (x3, y3), (x4, y4)]
        points.sort(key=lambda p: p[0])  # 按x坐标排序
        return [points[0], points[-1]]  # 返回x坐标最小和最大的点

    def _merge_segments(self,segments):
        if not segments:
            return []
        
        result = segments.copy()
        i = 0
        
        while i < len(result):
            j = i + 1
            merged = False
            while j < len(result):
                if self._segments_overlap(result[i], result[j]):
                    # 合并重叠的线段
                    merged_segment = self._merge_two_segments(result[i], result[j])
                    result[i] = merged_segment
                    result.pop(j)
                    merged = True
                else:
                    j += 1
            if not merged:
                i += 1
        
        return result


    def _extend_segments(self, segments, extend_length):
        """
        将首尾相接的线段列表的起始两端延长指定长度
        参数:
        segments: 线段列表，每个线段包含起点和终点坐标 [(x1,y1), (x2,y2)]，线段首尾相接
        extend_length: 需要延长的长度

        返回:
        extended_segments: 延长后的线段列表
        """
        if not segments or len(segments) < 1:
            return segments

        extended_segments = segments.copy()  # 复制原始线段列表
        for i in [0,-1]:
            start_point = np.array(extended_segments[i][0])
            end_point = np.array(extended_segments[i][1])
            # 计算方向向量
            direction = end_point - start_point
            if i == 0:
                #延长起点
                extended_segments[i][0] = self._extend_point(start_point, direction, -extend_length)
            else:
                #延长终点
                extended_segments[i][1] = self._extend_point(end_point, direction, extend_length)
        return extended_segments


    def _get_intersection(self,seg1: Segment, seg2: Segment) -> Tuple[float, float]:
        """计算两条线段的交点，使用numpy加速计算"""
        if not seg1.bbox.overlaps(seg2.bbox):
            return None

        p1, p2 = seg1.start, seg1.end
        p3, p4 = seg2.start, seg2.end

        denominator = np.cross(p2 - p1, p4 - p3)
        if denominator == 0:  # 平行或重合
            return None

        t = np.cross(p3 - p1, p4 - p3) / denominator
        u = np.cross(p3 - p1, p2 - p1) / denominator

        if 0 <= t <= 1 and 0 <= u <= 1:
            return tuple(p1 + t * (p2 - p1))
        return None


    def _split_segments_optimized(
        self,
        segments_data: List[List[Tuple[float, float]]]
    ) -> List[List[Tuple[float, float]]]:
        """使用扫描线算法处理线段交点"""
        # 转换输入数据为Segment对象
        segments = [Segment(start, end) for start, end in segments_data]

        # 创建事件点列表（线段的起点和终点）
        events = []
        for i, seg in enumerate(segments):
            x_min = min(seg.start[0], seg.end[0])
            x_max = max(seg.start[0], seg.end[0])
            events.append((x_min, 'start', i))
            events.append((x_max, 'end', i))

        # 按x坐标排序事件点
        events.sort()

        # 存储活动线段（当前扫描线相交的线段）
        active_segments = set()
        # 存储需要分割的线段和交点
        intersections = []

        # 扫描线算法
        for x, event_type, seg_idx in events:
            if event_type == 'start':
                # 检查新线段是否与活动线段相交
                for active_idx in active_segments:
                    intersection = self._get_intersection(segments[seg_idx], segments[active_idx])
                    if intersection:
                        intersections.append((seg_idx, active_idx, intersection))
                active_segments.add(seg_idx)
            else:
                try:
                    active_segments.remove(seg_idx)
                except KeyError:
                    continue

        # 处理所有交点
        result = segments_data.copy()
        processed_pairs = set()

        for seg_idx1, seg_idx2, intersection in intersections:
            if (seg_idx1, seg_idx2) in processed_pairs:
                continue

            # 获取原始线段
            seg1 = result[seg_idx1]
            seg2 = result[seg_idx2]

            # 创建新的分割线段
            new_segments = [
                [seg1[0], intersection],
                [intersection, seg1[1]],
                [seg2[0], intersection],
                [intersection, seg2[1]]
            ]

            # 更新结果
            result[seg_idx1] = new_segments[0]
            result[seg_idx2] = new_segments[2]
            result.extend([new_segments[1], new_segments[3]])

            processed_pairs.add((seg_idx1, seg_idx2))
            processed_pairs.add((seg_idx2, seg_idx1))

        return result


    def _check_segment_connections(self, segment, all_segments, threshold=0.1):
        """
        检查线段的两个端点是否都至少与其他线段相连
        通过计算端点距离判断连接，距离小于threshold即认为连接
        返回True表示线段应该保留，False表示应该删除
        """
        p1, p2 = np.array(segment[0]), np.array(segment[1])
        p1_connected = False
        p2_connected = False

        for other in all_segments:
            if other == segment:
                continue

            other_p1 = np.array(other[0])
            other_p2 = np.array(other[1])

            # 检查第一个端点是否与其他线段相连
            if np.linalg.norm(p1 - other_p1) < threshold or np.linalg.norm(p1 - other_p2) < threshold:
                p1_connected = True

            # 检查第二个端点是否与其他线段相连
            if np.linalg.norm(p2 - other_p1) < threshold or np.linalg.norm(p2 - other_p2) < threshold:
                p2_connected = True

            # 如果两个端点都已找到连接，可以提前结束搜索
            if p1_connected and p2_connected:
                break

        # 只有当两个端点都有连接时才返回True
        return p1_connected and p2_connected


    def _remove_isolated_segments(self, segments):
        """
        迭代删除不符合要求的线段（任一端点未与其他线段相连）

        Args:
            segments: 线段列表，每个线段格式为 [(x1,y1), (x2,y2)]

        Returns:
            删除不合格线段后的线段列表
        """
        if not segments:
            return []

        result = segments.copy()
        changed = True

        while changed:
            changed = False
            segments_to_remove = []

            # 检查每条线段
            for segment in result:
                # 如果线段的任一端点未连接，则标记为删除
                if not self._check_segment_connections(segment, result):
                    segments_to_remove.append(segment)
                    changed = True

            # 删除标记的线段
            for segment in segments_to_remove:
                result.remove(segment)

        return result


    def _is_same_segment(self,seg1, seg2):
        """检查两个线段是否相同（考虑正向和反向）"""
        # 正向比较
        forward_same = (seg1[0] == seg2[0] and seg1[1]==seg2[1])
        # 反向比较
        reverse_same = (seg1[0] == seg2[1] and seg1[1]==seg2[0])
        return forward_same or reverse_same


    def _remove_duplicate_segments(self, segments):
        """
        删除重复的线段，包括：
        1. 首尾坐标相同的线段
        2. 首尾翻转后与其他线段相同的线段

        参数:
        segments: 线段列表，每个线段是一个包含两个坐标点的列表
        tolerance: 浮点数比较的容差值

        返回:
        unique_segments: 去重后的线段列表
        """
        # 存储要保留的线段
        unique_segments = []
        removed_indices = set()

        # 遍历所有线段
        for i, segment in enumerate(segments):
            # 如果已经被标记为删除，则跳过
            if i in removed_indices:
                continue

            # 检查是否是首尾相同的线段
            if segment[0]== segment[1]:
                removed_indices.add(i)
                continue

            # 检查是否与之前的线段重复
            is_duplicate = False
            for j in range(len(unique_segments)):
                if self._is_same_segment(segment, unique_segments[j]):
                    is_duplicate = True
                    break

            if not is_duplicate:
                unique_segments.append(segment)

        return unique_segments


    def visualize_segments_and_polygons(self,original_segments, result_segments, polygon):
        """
        可视化原始线段、处理后的线段和生成的多边形
        """
        # 创建三个子图
        fig, (ax1, ax2, ax3) = plt.subplots(1, 3, figsize=(10, 7))

        # 设置标题
        ax1.set_title('Original Segments')
        ax2.set_title('Processed Segments')
        ax3.set_title('Generated Polygons')

        # 获取所有点的坐标范围
        all_points = np.array([point for segment in original_segments for point in segment])
        x_min, y_min = np.min(all_points, axis=0)
        x_max, y_max = np.max(all_points, axis=0)

        # 添加边距
        margin = (x_max - x_min) * 0.1

        # 设置坐标轴范围
        for ax in [ax1, ax2, ax3]:
            ax.set_xlim(x_min - margin, x_max + margin)
            ax.set_ylim(y_min - margin, y_max + margin)
            ax.grid(True, linestyle='--', alpha=0.7)
            ax.set_aspect('equal')

        # 绘制原始线段
        for segment in original_segments:
            x_coords = [segment[0][0], segment[1][0]]
            y_coords = [segment[0][1], segment[1][1]]
            ax1.plot(x_coords, y_coords, 'b-', linewidth=1.5, alpha=0.7)
            ax1.plot(x_coords, y_coords, 'r.', markersize=8)

        # 绘制处理后的线段
        colors = plt.cm.rainbow(np.linspace(0, 1, len(result_segments)))
        for segment, color in zip(result_segments, colors):
            x_coords = [segment[0][0], segment[1][0]]
            y_coords = [segment[0][1], segment[1][1]]
            ax2.plot(x_coords, y_coords, color=color, linewidth=1.5, alpha=0.7)
            ax2.plot(x_coords, y_coords, 'r.', markersize=8)

        # 绘制多边形
        polygon_color = plt.cm.Pastel1(0)
        x, y = polygon.exterior.xy
        ax3.fill(x, y, alpha=0.5, fc=polygon_color, ec='black')

        # 添加标签
        for ax in [ax1, ax2, ax3]:
            ax.set_xlabel('X')
            ax.set_ylabel('Y')

        plt.tight_layout()
        plt.show()


    def _segments_to_polygons(self, segments):
        """
        将线段列表转换为多边形
        segments: 线段列表，每个线段包含起点和终点坐标
        返回: 多边形列表
        """
        # 将线段转换为LineString对象
        lines = [LineString(segment) for segment in segments]
        # 创建MultiLineString
        # multi_line = MultiLineString(lines)
        # print(multi_line)
        # 使用polygonize函数生成多边形
        polygons = list(polygonize(lines))
        return polygons


    def create_polygon(self, segments_datas, cfg):
        #输入线段列表，输出多边形
        try:
            # 扩展线段
            origin_seg = [segment for segments in segments_datas for segment in segments]
            extend_length = 2
            extended = [self._extend_segments(segments_data, extend_length) for segments_data in segments_datas]
            # 将extended转化为1维列表
            extended = [segment for segments in extended for segment in segments]
            # 合并线段
            extended = self._merge_segments(extended)
        except Exception as e:
            print(f"边界识别错误(线段拓展出错): {e}")
            return None,None
        try:
            #用交点打断线段
            result = self._split_segments_optimized(extended)
            #删去旁支
            result1 = self._remove_isolated_segments(result)
            #去重
            result1 = self._remove_duplicate_segments(result1)
        except Exception as e:
            print(f"边界识别错误(边间简化出错): {e}")
            return None,None
        if cfg.is_visual:
            self.visualize_segments_and_polygons(extended,result1, Polygon([(0,0),(0,100),(100,100),(100,0)]))
        try:
            #转化为多边形
            polygon = self._segments_to_polygons(result1)[0]
            return extended,polygon
        except Exception as e:
            print(f"边界识别错误（边界融合出错）: {e}")
            return None,None


    def ask_user_to_continue(self):
        """
        在命令行中询问用户是否退出

        Args:
            cad_app: CAD 应用程序对象 (例如, win32com.client.Dispatch("AutoCAD.Application") 返回的对象).

        Returns:
            bool: 如果用户输入 "y" 或 "Y", 则返回 True; 否则返回 False.
        """
        try:
            while True:
                answer = self._doc.Utility.GetString(1,"是否继续当前程序？[退出程序(N)/继续(Y)] <继续>: ")
                if answer.upper().startswith("Y") or answer == "":  # 空字符串表示用户直接按了 Enter
                    return True
                elif answer.upper().startswith("N"):
                    return False
                else:
                    self._doc.Utility.Prompt("无效输入。 请输入 'y' 或 'n'。")  # 提示用户输入有效值
        except Exception as e:
            # print(f"发生错误: {e}")
            return False  # 发生错误时，退出


    def get_area_input(self):
        """
        获取用户输入的面积值，可以通过选择包含面积的文本或手动输入。

        Returns:
            str or float: 用户输入的面积值或选择的文本内容，如果出错则返回 None。
        """
        while True: # 循环直到获取有效的数值输入或用户取消
            try:
                area_str = self._doc.Utility.GetString(1, "\n输入面积: ")
                area = float(area_str)
                return area
            except ValueError:
                self._doc.Utility.Prompt("输入的面积值无效。\n")
            except: # 用户按了取消
                return None


    def _find_cutting_line_with_slope(self, polygon, target_area, slope=0, initial_y_bounds=None, tolerance=0.00001, max_iterations=1000):
        """
        找到一条具有给定斜率的直线，将多边形分割为两部分，使得下部多边形的面积等于目标面积。
        通过迭代扩展搜索区间来处理f(a)和f(b)同号的情况。
        """

        def area_difference(y_intercept):
            """计算下部多边形面积与目标面积之间的差值。
            """
            min_x, min_y, max_x, max_y = polygon.bounds

            # Calculate line points at polygon bounds
            x1 = min_x
            y1 = y_intercept + slope * x1
            x2 = max_x
            y2 = y_intercept + slope * x2

            line = LineString([(x1, y1), (x2, y2)])

            # Create a polygon representing the area below the line
            lower_polygon_points = [
                (min_x, min_y),
                (max_x, min_y),
                (x2, y2),
                (x1, y1)
            ]
            lower_polygon = polygon.intersection(Polygon(lower_polygon_points))

            return lower_polygon.area - target_area

        min_x, min_y, max_x, max_y = polygon.bounds
        if initial_y_bounds is None:
            y_min = min_y
            y_max = max_y
        else:
            y_min, y_max = initial_y_bounds
            y_min = max(y_min, min_y)
            y_max = min(y_max, max_y)


        # --- 改进符号检查和区间扩展 ---
        # 计算区间下限处的面积差值
        fa = area_difference(y_min)
        # 计算区间上限处的面积差值
        fb = area_difference(y_max)

        # 限制扩展尝试次数
        for _ in range(max_iterations):
            # 如果 fa 和 fb 的符号不同，则跳出循环，进行二分查找
            if fa * fb <= 0:
                break

            # 向绝对值较大的方向扩展区间
            if abs(fa) > abs(fb):
                # 向下扩展区间
                y_min -= (y_max - y_min)
                # 确保 y_min 在多边形的边界内
                y_min = max(y_min, min_y)
                # 重新计算扩展后区间下限处的面积差值
                fa = area_difference(y_min)
            else:
                # 向上扩展区间
                y_max += (y_max - y_min)
                # 确保 y_max 在多边形的边界内
                y_max = min(y_max, max_y)
                # 重新计算扩展后区间上限处的面积差值
                fb = area_difference(y_max)
        if fa * fb > 0:
            # 区间扩展未能找到符号相反的情况
            # print("错误：未能找到一个有效的区间，使得 f(a) 和 f(b) 的符号不同。")
            return None

        try:
            result = bisect(area_difference, y_min, y_max, xtol=tolerance)
        except ValueError as e:
            print(f"Bisection failed: {e}")
            return None

        return result


    def get_intersection_points(self, polygon, line):
        """计算并返回直线与多边形的交点"""
        intersections = polygon.boundary.intersection(line)

        if intersections.is_empty:
            return []
        elif intersections.geom_type == 'Point':
            return [(intersections.x, intersections.y)]
        elif intersections.geom_type == 'MultiPoint':
            # print("直线与边界相交于多个点，只返回两个点")
            points = [(p.x, p.y) for p in intersections.geoms]
            return [points[0], points[-1]]
        else:
            return []


    def process_polygon(self, polygon, target_area, slope=0.0):
        """
        处理多边形分割
        :param polygon_coords: 多边形坐标点列表
        :param target_area: 目标面积
        :param slope: 分割线斜率，默认为0(水平线)
        :param visualize: 是否可视化结果，默认为False
        :return: (cutting_line, lower_polygon) 分割线和底部多边形
        """
        y_intercept = self._find_cutting_line_with_slope(polygon, target_area, slope)

        if y_intercept is None:
            return None, None

        min_x, min_y, max_x, max_y = polygon.bounds
        x1 = min_x
        y1 = y_intercept + slope * x1
        x2 = max_x
        y2 = y_intercept + slope * x2

        # 更新交点坐标为起终点坐标
        intersections = self.get_intersection_points(polygon, LineString([(x1, y1), (x2, y2)]))
        cutting_line = LineString(intersections) if intersections else LineString([(x1, y1), (x2, y2)])

        # 创建底部多边形
        lower_polygon_points = [
            (min_x, min_y),
            (max_x, min_y),
            (x2, y2),
            (x1, y1)
        ]
        lower_polygon = polygon.intersection(Polygon(lower_polygon_points))

        return cutting_line, lower_polygon


    def _vtpnt(self,x, y, z=0):
        return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (x, y, z))

    def _vtobj(self,obj):
        return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, obj)

    def _vtfloat(self,lst):
        return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, lst)


    def create_polyline(self, points, color=4):
        """
        在 AutoCAD 中创建多段线。

        Args:
            points: 多段线的坐标点序列，例如 [(x1, y1), (x2, y2), ...]
        """
        for i in range(5):  # 最多尝试5次
            try:
            # 将坐标点转换为 AutoCAD 可接受的格式 (VARIANT)
                coords = []
                for point in points:
                    coords.append(point[0])
                    coords.append(point[1])
                    # coords.append(0)  # 可选的 Z 坐标，这里设为 0
                points = self._vtfloat(coords)
                plineObj = self._msp.AddLightWeightPolyline(points)
                # 指定颜色
                plineObj.Color = color
                self._doc.Application.Update()
                return plineObj # 返回创建的多段线对象
            except Exception as e:
                print(f"创建多段线失败，第{i+1}次尝试: {e}")
                time.sleep(1)
                pass
        print("创建多段线失败",points)
        return None


    def fill_polygon(self, points, hatch_pattern="GRAVEL", color=4,scale = 1.0):
        """
        在 AutoCAD 中填充多边形。

        Args:
            points: 多边形的坐标点序列，例如 [(x1, y1), (x2, y2),...]
            hatch_pattern: 填充图案的名称，默认为 "ANSI31"
            color: 填充颜色的索引，默认为 4 (蓝色)
            scale: 缩放比例，默认为 1.0
        """
        try:
            plineObj = self.create_polyline(points, color=7)
            if plineObj is None:
                raise Exception("创建填充区边界失败")
            for i in range(5):  # 最多尝试5次
                try:
                    plineObj.Closed = True
                    outerLoop = []
                    outerLoop.append(plineObj)
                    outerLoop = self._vtobj(outerLoop)
                    hatchObj = self._doc.ModelSpace.AddHatch(0, hatch_pattern, True)
                    hatchObj.AppendOuterLoop(outerLoop)
                    hatchObj.PatternScale = scale  # 设置填充比例
                    hatchObj.Evaluate()  # 进行填充计算，使图案吻合于边界。
                    return hatchObj
                except Exception as e:
                    print(f"填充失败，第{i+1}次尝试: {e}")
                    time.sleep(1)  # 等待1秒后重试
        except Exception as e:
            print(f"填充多边形失败: {e}")
