from readers import MeterIDReader, DeviationReader

class APDeviationFormFiller:
    def __init__(self, meter_address, reverse):
        self._meter_address = meter_address
        self._meter_id = MeterIDReader(meter_address).read()
        power_type = '0' if not reverse else '1'
        self._deviation_reader_banlanced = DeviationReader(self._meter_id, power_type, 'balanced')
        self._deviation_reader_unbanlanced = DeviationReader(self._meter_id, power_type, 'unbalanced')

    def fill(self):
        pass