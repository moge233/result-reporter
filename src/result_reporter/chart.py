#! python3


from pydrf.textchart import Header, RaceData, StarterPerformanceData

from .race import Race


class Chart:
    def __init__(self, header: Header, races: list[RaceData], starters: list[StarterPerformanceData]):
        self.track_code: str = header.track_code
        self.race_date: str = header.race_date
        self.number_of_races: int = header.number_of_races
        self.races: list[Race] = []
        for race in races:
            data = race
            race_number = race.race_number
            horses: list[StarterPerformanceData] = []
            for starter in starters:
                if starter.race_number == race_number:
                    horses.append(starter)
            self.races.append(Race(data, horses))

    def __str__(self):
        ret = ''
        for k, v in vars(self).items():
            ret += f'{k}={v}, '
        return f'Chart({ret[:-2]})'

    def __repr__(self):
        ret = ''
        for k, v in vars(self).items():
            ret += f'{k}={v}, '
        return f'Chart({ret[:-2]})'
