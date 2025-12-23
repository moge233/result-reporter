#! python3

from abc import ABC
from math import floor, nan


FEET_PER_BEATEN_LENGTH: float = 11.0
DEFAULT_MAXIMUM_SPRINT_DISTANCE: float = 7.5


def get_time_of_beaten_lengths(distance: float, final_time: float) -> float:
    beaten_lengths_in_race: float = (distance * 660) / FEET_PER_BEATEN_LENGTH
    return final_time / beaten_lengths_in_race


class Report(ABC):
    pass


class ShakeUpReport(Report):
    def __init__(self, key: str, cls: str, claiming_price: float, purse: float,
                 surface: str, distance: float, post_position: int,
                 bl1: float, bl2: float, bl3: float, blf: float,
                 fr1: float, fr2: float, fr3: float, finish: float):
        super().__init__()
        self.key: str = key
        self.cls: str = cls
        self.claiming_price: float = claiming_price
        self.purse: float = purse
        self.surface: str = surface
        self.post_position: int = post_position
        self.distance: float = round(distance / 100, 1)
        self.bl1: float = bl1 / 100
        self.bl2: float = bl2 / 100
        self.bl3: float = bl3 / 100
        self.blf: float = blf / 100
        self.fr1: float = floor((fr1 + get_time_of_beaten_lengths(self.distance, finish) * self.bl1) * 10) / 10
        self.fr2: float = floor((fr2 + get_time_of_beaten_lengths(self.distance, finish) * self.bl2) * 10) / 10
        if self.distance > 6:
            self.fr3: float = floor((fr3 + get_time_of_beaten_lengths(self.distance, finish) * self.bl3) * 10) / 10
        elif self.distance == 6:
            self.fr3: float = floor((finish + get_time_of_beaten_lengths(self.distance, finish) * self.blf) * 10) / 10
        else:
            self.fr3: float = nan
        self.finish: float = floor((finish + get_time_of_beaten_lengths(self.distance, finish) * self.blf) * 10) / 10

    def __str__(self):
        ret = ''
        for k, v in vars(self).items():
            ret += f'{k}={v}, '
        return f'ShakeUpReport({ret[:-2]})'

    def __repr__(self):
        ret = ''
        for k, v in vars(self).items():
            ret += f'{k}={v}, '
        return f'ShakeUpReport({ret[:-2]})'


class DailyShakeUpReport:
    def __init__(self, all_surfaces: list[str], reports: list[ShakeUpReport]):
        self.surfaces: list[str] = all_surfaces
        for report in reports:
            pass


class BrohamerReport(Report):
    def __init__(self, key: str, cls: str, sex: str, age: str, claiming_price: float, purse: float, race: int,
                 surface: str, course: str, distance: float, number: int, post: int,
                 bl1: float, bl2: float, c1: float, c2: float, fc: float, comment=None):
        super().__init__()
        self.key: str = key
        self.cls: str = cls
        self.sex: str = sex
        self.age: str = age
        self.claiming_price: float = claiming_price
        self.purse: float = purse
        self.race: int = race
        self.surface: str = surface
        self.course: str = course
        self.distance: float = round(distance / 100, 1)
        self.number: int = number
        self.post: int = post
        self.bl1: float = round(bl1 / 100, 2)
        self.bl2: float = round(bl2 / 100, 2)
        self.c1: float = c1
        self.c2: float = c2
        self.fc: float = fc
        if self.distance < 5.0:
            self.fr1 = nan
            self.fr2 = nan
            self.fr3 = nan
            self.ep = nan
            self.sp = nan
            self.ap = nan
            self.fx = nan
            self.energy = nan
        elif self.distance <= DEFAULT_MAXIMUM_SPRINT_DISTANCE:
            if not self.c1 or not self.c2:
                self.fr1 = nan
                self.fr2 = nan
                self.fr3 = nan
                self.ep = nan
                self.sp = nan
                self.ap = nan
                self.fx = nan
                self.energy = nan
            else:
                self.fr1 = round((1320 - 10 * self.bl1) / self.c1, 1)
                self.fr2 = round(
                    (1320 - 10 * (self.bl2 - self.bl1)) / (self.c2 - self.c1), 1)
                self.fr3 = round((660 * (self.distance - 4) + 10 * self.bl2) / (self.fc - self.c2), 1)
                self.ep = round((2640 - 10 * self.bl2) / self.c2, 1)
                self.sp = round((self.ep + self.fr3) / 2, 1)
                self.ap = round((self.fr1 + self.fr2 + self.fr3) / 3, 1)
                self.fx = round((self.fr1 + self.fr3) / 2, 1)
                self.energy = round(self.ep / (self.ep + self.fr3), 2)
        else:
            if not self.c1 or not self.c2:
                self.fr1 = nan
                self.fr2 = nan
                self.fr3 = nan
                self.ep = nan
                self.sp = nan
                self.ap = nan
                self.fx = nan
                self.energy = nan
            else:
                self.fr1 = round((2640 - 10 * self.bl1) / self.c1, 1)
                self.fr2 = round(
                    (1320 - 10 * (self.bl2 - self.bl1)) / (self.c2 - self.c1), 1)
                self.fr3 = round((660 * (self.distance - 6) + 10 * self.bl2) / (self.fc - self.c2), 1)
                self.ep = round((3960 - 10 * self.bl2) / self.c2, 1)
                self.sp = round((self.ep + self.fr3) / 2, 1)
                self.ap = round((self.fr1 + self.fr3) / 2, 1)
                self.fx = round((self.fr1 + self.fr3) / 2, 1)
                self.energy = round(self.ep / (self.ep + self.fr3), 1)
        self.comment = comment

    def __str__(self):
        ret = ''
        for k, v in vars(self).items():
            ret += f'{k}={v}, '
        return f'BrohamerReport({ret[:-2]})'

    def __repr__(self):
        ret = ''
        for k, v in vars(self).items():
            ret += f'{k}={v}, '
        return f'BrohamerReport({ret[:-2]})'


class DailyBrohamerReport:
    def __init__(self, all_surfaces: list[str], reports: list[ShakeUpReport]):
        self.surfaces: list[str] = all_surfaces
        for report in reports:
            pass
