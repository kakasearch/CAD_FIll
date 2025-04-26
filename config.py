import os
import json

class Config:
    def __init__(self, config_path: str):
        if not os.path.exists(config_path):
            print('【警告】未找到config.json配置文件，使用默认配置')
            self.is_visual = True
            self.fill_pattern = 'ANGLE',
            self.fill_scale = 0.1,
            self.fill_color = 7
            return

        with open(config_path, 'r', encoding='utf-8') as file:
            data = json.load(file)

        self.is_visual = data.get('is_visual', True)
        self.fill_pattern = data.get('fill_pattern', 'ANGLE')
        self.fill_scale = data.get('fill_scale', 0.1)
        self.fill_color = data.get('fill_color', 7)
