import os
import random
from config import Config
from cad_app import CadApp

def main_process():
    # 读取配置文件
    cfg = Config('./config.json')
    # 获取 AutoCAD 应用
    app = CadApp()
    if not app.load():
        print("【警告】无法连接到 AutoCAD!!")
        os.system("pause")
        return "CAD连接失败，程序已终止。"

    print("【提示】AutoCAD 连接成功！\n【提示】请在cad中选择边界")
    while True:
        try:
            segments_datas = app.get_boundary_coords("\n请选择边界")
            if not segments_datas:
                print("【提示】未选择边界")
                if not app.ask_user_to_continue():
                    return "未选择边界，程序已终止。"

            original_segments, polygon =  app.create_polygon(segments_datas, cfg)
            #开始分割，寻找符合要求的直线，将多边形分割为两个部分，其中下半部分的面积等于目标面积
            if not polygon:
                print("【警告】未找到边界！")
                continue
            target_area = app.get_area_input() # 目标面积
            slope = round(random.uniform(-0.0002, 0.0002), 5) #分割线斜率，正负1%的随机值
            print(f"【提示】目标面积为：{target_area}，随机分割线斜率为：{slope}")
            if target_area > polygon.area:
                print("【警告】目标面积大于断面面积!!")
                continue
            cutting_line, lower_polygon = app.process_polygon(polygon, target_area,slope=slope)

            # 在 AutoCAD 中创建多段线
            if not cutting_line:
                print("【警告】无法找到合适的分割线!!")
                continue
            #填充图形
            polygon_coords = list(lower_polygon.exterior.coords)

            app.fill_polygon(
                polygon_coords + polygon_coords[:1],
                hatch_pattern= cfg.get("fill_pattern"),
                color= cfg.get("fill_color"),
                scale = cfg.get("fill_scale")
            )
            polyline = app.create_polyline(cutting_line.coords) # 将所有边界合并为一个多段线
            if polyline:
                print("【提示】顶部边界线已在 AutoCAD 中创建。")
            else:
                print("【提示】顶部边界线创建失败。")
            # 可视化分割结果
            # app.plot_polygon_with_line(polygon, cutting_line, lower_polygon)
            # 将分割线转化为线段
            # 可视化结果
            if cfg.is_visual:
                cutting_line_point = list(cutting_line.coords)
                cutting_line_segments = []
                for i in range(len(cutting_line_point)-1):
                    cutting_line_segments.append([cutting_line_point[i], cutting_line_point[i+1]])
                    app.visualize_segments_and_polygons(
                        original_segments,
                        original_segments + cutting_line_segments,
                        lower_polygon
                    )
            if not app.ask_user_to_continue():
                return "程序已终止。"
        except Exception as e:
            print(f"【提示】程序发生错误: {e}")
        finally:
            print("\n\n")

# 使用示例
if __name__ == "__main__":
    os.system("title cad边界识别 by kaka")
    while True:
        main_process()
        is_exit = input("【提示】程序已暂停，按任意键继续，输入y退出：")
        if is_exit.lower() == "y":
            break
