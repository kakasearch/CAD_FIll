import numpy as np
from dataclasses import dataclass
from typing import Tuple

@dataclass
class BoundingBox:
    x_min: float
    y_min: float
    x_max: float
    y_max: float

    def overlaps(self, other: 'BoundingBox') -> bool:
        return not (self.x_max < other.x_min or
                   other.x_max < self.x_min or
                   self.y_max < other.y_min or
                   other.y_max < self.y_min)

class Segment:
    def __init__(self, start: Tuple[float, float], end: Tuple[float, float]):
        self.start = np.array(start)
        self.end = np.array(end)
        self.bbox = self._compute_bbox()


    def _compute_bbox(self) -> BoundingBox:
        x_min = min(self.start[0], self.end[0])
        x_max = max(self.start[0], self.end[0])
        y_min = min(self.start[1], self.end[1])
        y_max = max(self.start[1], self.end[1])
        return BoundingBox(x_min, y_min, x_max, y_max)
