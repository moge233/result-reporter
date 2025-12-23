#! python3


from pydrf.textchart import RaceData, StarterPerformanceData


class Race:
    def __init__(self, data: RaceData, starters: list[StarterPerformanceData]):
        self.data: RaceData = data
        self.starters: list[StarterPerformanceData] = starters

    def __str__(self):
        ret = ''
        for k, v in vars(self).items():
            ret += f'{k}={v}, '
        return f'Race({ret[:-2]})'

    def __repr__(self):
        ret = ''
        for k, v in vars(self).items():
            ret += f'{k}={v}, '
        return f'Race({ret[:-2]})'
