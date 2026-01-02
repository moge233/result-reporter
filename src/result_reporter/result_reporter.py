#! python3


import csv
import os

from pydrf.textchart import Header, RaceData, RecordType, StarterPerformanceData

from src.result_reporter.chart import Chart
from src.result_reporter.coursetype import CourseType


class ResultReporter:
    """
    Main object of the result reporter application.
    """
    def __init__(self, track_code: str, path: str):
        self._track_code: str = track_code
        self._charts_path: str = path
        self._charts: list[Chart] = self._get_charts()
        self._course_types: list[CourseType] = self._get_course_types()

    @property
    def track_code(self) -> str:
        """
        track_code get this ResultReporter instance's current track code

        Returns:
            str: current track code in use
        """
        return self._track_code

    def set_track_code(self, track_code: str) -> None:
        """
        set_track_code set this ResultReporter instance's track code

        Note: this method also updates the instance's charts and course types to reflect
        the new track code

        Args:
            track_code (str): this ResultReporter instance's new track code
        """
        self._track_code = track_code
        self._charts = self._get_charts()
        self._course_types = self._get_course_types()

    @property
    def charts_path(self) -> str:
        """
        charts_path get this ResultReporter instance's current charts path

        Returns:
            str: current charts path in use
        """
        return self._charts_path

    def set_charts_path(self, charts_path: str) -> None:
        """
        set_charts_path set this ResultReporter instance's charts path

        Note: this method also updates the instance's charts and course types to reflect
        the charts in the new charts path

        Args:
            charts_path (str): this ResultReporter instance's new charts path
        """
        self._charts_path = charts_path
        self._charts = self._get_charts()
        self._course_types = self._get_course_types()

    @property
    def charts(self) -> list[Chart]:
        """
        charts get this isntance's list of Chart objects

        Returns:
            list[Chart]: a list of Chart objects
        """
        return self._charts

    @property
    def course_types(self) -> list[CourseType]:
        """
        course_types get this instance's list of CourseType objects

        Returns:
            list[CourseType]: a list of CourseType objects
        """
        return self._course_types

    def _parse_chart(self, path: str) -> Chart | None:
        '''
        parse a DRF text chart file into a Chart type

        :param self: instance of ResultReporter
        :param path: absolute path to the DRF text chart file
        :type path: str
        :return: return the encapsulated Chart object for this DRF text chart file
        :rtype: Chart | None
        '''
        header: Header | None = None
        race_data: list[RaceData] = []
        starters_performance_data: list[StarterPerformanceData] = []
        try:
            with open(path) as chart_file:
                reader = csv.reader(chart_file.readlines())
                for line in reader:
                    if line[0] == RecordType.HEADER:
                        header = Header.create(line)
                    elif line[0] == RecordType.RACE:
                        race_data.append(RaceData.create(line))
                    elif line[0] == RecordType.STARTER:
                        starters_performance_data.append(StarterPerformanceData.create(line))
                    elif line[0] == RecordType.EXOTIC_WAGERING:
                        pass
                    elif line[0] == RecordType.ATTENDANCE:
                        pass
                    elif line[0] == RecordType.COMMENT:
                        pass
                    elif line[0] == RecordType.FOOTNOTE:
                        pass
            if header and race_data and starters_performance_data:
                return Chart(
                    header,
                    race_data,
                    starters_performance_data
                )
            return None
        except FileNotFoundError as e:
            print(f'{e}: could not find result file: {path}')
            return None

    def _get_charts(self) -> list[Chart]:
        charts: list[Chart] = []
        for dir in os.listdir(self._charts_path):
            dir_path = os.path.join(self._charts_path, dir)
            for chart_path in os.listdir(dir_path):
                if chart_path[:len(self._track_code)] == self._track_code and \
                        chart_path[len(self._track_code)].isdigit():
                    chart_path = os.path.join(dir_path, chart_path)
                    chart: Chart | None = self._parse_chart(chart_path)
                    if chart:
                        charts.append(chart)
        return charts

    def _get_surfaces(self) -> list[str]:
        """
        get_surfaces get a list of surfaces for the ResultReporter instance's track code

        Returns:
            list[str]: a list of DRF surface designators
        """
        surfaces: list[str] = []
        for chart in self._charts:
            for race in chart.races:
                if race.data.course_type not in surfaces:
                    surfaces.append(race.data.course_type)
        return sorted(surfaces)

    def _get_course_types(self) -> list[CourseType]:
        """
        get_course_types get a list of CourseTypes for the ResultReporter instance's track code

        Returns:
            list[CourseType]: a list of CourseTypes
        """
        course_types: list[CourseType] = []
        for chart in self._charts:
            for race in chart.races:
                course_type: CourseType = CourseType.parse_course_type(race.data.course_type)
                if course_type not in course_types:
                    course_types.append(course_type)
        return course_types
