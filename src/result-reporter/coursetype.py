#!python3


from enum import Enum


class CourseType(Enum):
    ALL_WEATHER_TRAINING = 1
    DIRT = 2
    TURF = 3
    INNER_TURF = 4
    OUTER_TURF = 5
    ALL_WEATHER_TRACK = 6
    DIRT_TRAINING = 7
    INNER_TRACK = 8
    WOOD_CHIPS = 9
    TIMBER = 10
    DOWNHILL_TURF = 11
    TURF_TRAINING = 12
    JUMP = 13
    HURDLE = 14
    STEEPLECHASE = 15
    HUNT_ON_TURF = 16

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

    @staticmethod
    def parse_course_type(course: str) -> 'CourseType':
        if course == 'A':
            return CourseType.ALL_WEATHER_TRAINING
        elif course == 'D':
            return CourseType.DIRT
        elif course == 'E':
            return CourseType.ALL_WEATHER_TRACK
        elif course == 'F':
            return CourseType.DIRT_TRAINING
        elif course == 'N':
            return CourseType.INNER_TRACK
        elif course == 'W':
            return CourseType.WOOD_CHIPS
        elif course == 'B':
            return CourseType.TIMBER
        elif course == 'C':
            return CourseType.DOWNHILL_TURF
        elif course == 'G':
            return CourseType.TURF_TRAINING
        elif course == 'I':
            return CourseType.INNER_TURF
        elif course == 'J':
            return CourseType.JUMP
        elif course == 'M':
            return CourseType.HURDLE
        elif course == 'O':
            return CourseType.OUTER_TURF
        elif course == 'S':
            return CourseType.STEEPLECHASE
        elif course == 'T':
            return CourseType.TURF
        elif course == 'U':
            return CourseType.HUNT_ON_TURF
        # Default to dirt
        return CourseType.DIRT

    def course_to_str(self) -> str:
        if self is CourseType.ALL_WEATHER_TRACK:
            return 'Tapeta'
        elif self is CourseType.DIRT:
            return 'Dirt'
        elif self is CourseType.HURDLE:
            return 'Hurdle'
        elif self is CourseType.INNER_TRACK:
            return 'Inner Track'
        elif self is CourseType.INNER_TURF:
            return 'Inner Turf'
        elif self is CourseType.OUTER_TURF:
            return 'Outer Turf'
        elif self is CourseType.TURF:
            return 'Turf'
        else:
            return 'Dirt'  # as the default

    def course_to_letter(self) -> str:
        if self is CourseType.ALL_WEATHER_TRACK:
            return 'E'
        elif self is CourseType.DIRT:
            return 'D'
        elif self is CourseType.HURDLE:
            return 'M'
        elif self is CourseType.INNER_TRACK:
            return 'W'
        elif self is CourseType.INNER_TURF:
            return 'J'
        elif self is CourseType.OUTER_TURF:
            return 'O'
        elif self is CourseType.TURF:
            return 'T'
        else:
            return 'D'  # as the default
