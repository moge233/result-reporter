#! python3


from enum import Enum


class RaceType(Enum):
    SPRINT = 1
    ROUTE = 2

    def __eq__(self, other):
        if self.__class__ is other.__class__:
            return self.value == other.value
        return NotImplemented

    def __gt__(self, other):
        if self.__class__ is other.__class__:
            return self.value > other.value
        return NotImplemented

    def __lt__(self, other):
        if self.__class__ is other.__class__:
            return self.value < other.value
        return NotImplemented

    def __gte__(self, other):
        if self.__class__ is other.__class__:
            return self.value >= other.value
        return NotImplemented

    def __lte__(self, other):
        if self.__class__ is other.__class__:
            return self.value <= other.value
        return NotImplemented
