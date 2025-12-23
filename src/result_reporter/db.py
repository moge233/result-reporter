#! python3


from math import nan
from mssql_python.exceptions import IntegrityError

from mssql_python import connect, Connection, Cursor

from coursetype import CourseType
from racetype import RaceType


class ResultDatabaseManager:
    def __init__(self, sql_connection_string: str):
        self.connection: Connection = connect(sql_connection_string)

    def create_table_if_not_exists(self, table_name: str, columns: list[tuple[str, str]]) -> None:
        execution_string: str = f'IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE '\
                                f'TABLE_NAME = \'{table_name}\')'\
                                f'BEGIN CREATE TABLE {table_name}('
        execution_string += 'DATE DATE PRIMARY KEY,'
        for column in columns:
            name: str = column[0]
            datatype: str = column[1]
            execution_string += f'{name} {datatype},'
        execution_string = execution_string[:-1]
        execution_string += ') END'
        cursor: Cursor = self.connection.cursor()
        cursor.execute(execution_string)
        self.connection.commit()
        cursor.close()

    def get_column_names(self, table_name: str) -> list[str]:
        column_names: list[str] = []
        query_string: str = f'SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE '\
                            f'TABLE_NAME = \'{table_name}\' ORDER BY ORDINAL_POSITION;'
        cursor: Cursor = self.connection.cursor()
        cursor.execute(query_string)
        rows = cursor.fetchall()
        for row in rows:
            column_names.append(row[0])
        cursor.close()
        return column_names

    def add_record(self, table_name: str, values: list[tuple[float, float] | str]) -> None:
        '''
        Add a row to the given database.

        values are passed in as a list in the following format:
            date_str, (min_fr1, max_fr1), (min_fr2, max_fr2), (min_fr3, max_fr3), comment, ...
        so the length of values should be equal to (len(column_names) - 2) / 2
        '''
        column_names: list[str] = self.get_column_names(table_name)
        assert ((len(values) - 1) / 8) == ((len(column_names) - 1) / 14)
        n_surfaces: float = (len(values) - 1) / 8
        assert (n_surfaces % 1) == 0
        execution_string: str = f'INSERT INTO {table_name} ('
        for column_name in column_names:
            execution_string += f'{column_name},'
        execution_string = execution_string[:-1] + ') VALUES ('
        count: int = int(n_surfaces)
        row_str: str = '?,'
        row_vals: list[str | float] = [f'{values[0]}']
        for i in range(1, count * 8 + 1):
            if type(values[i]) is str:
                row_vals.append(values[i])  # type: ignore
                row_str += '?,'
            elif type(values[i]) is tuple:
                row_vals.append(values[i][0])
                row_vals.append(values[i][1])
                row_str += '?,?,'
        execution_string += row_str[:-1] + ');'
        clean_row_vals: list[str | float] = [0.0 if val is nan else val for val in row_vals]
        cursor: Cursor = self.connection.cursor()
        try:
            cursor.execute(execution_string, clean_row_vals)
        except IntegrityError:
            pass
        self.connection.commit()
        cursor.close()

    def get_record(self, table_name: str, race_date: str, course: CourseType, race_type: RaceType) -> list[float | str]:
        ret: list[float | str] = []
        query_string: str = f'SELECT * FROM {table_name} WHERE DATE = ?'
        cursor: Cursor = self.connection.cursor()
        cursor.execute(query_string, [race_date])
        fetch_result = cursor.fetchall()
        if cursor.description:
            column_names: list[str] = [
                description[0] for description in cursor.description
            ]
            results: list[str | float] = [
                result for result in fetch_result[0]
            ]
            course_str: str = course.course_to_str().upper()
            temp = list(zip(column_names, results))
            ret_list = [
                val[1] for val in temp if course_str in val[0]
            ]
            cursor.close()
        if ret_list:
            if race_type is RaceType.SPRINT:
                ret = ret_list[:7]
            elif race_type is RaceType.ROUTE:
                ret = ret_list[7:]  # the note is included
            if type(ret[-1]) is str:
                ret[-1] = ret[-1].rstrip()
        return ret
