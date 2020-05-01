from docx import Document


class Writer:
    def __init__(self, file_path):
        self._doc = Document(file_path)
        self._tables = self._doc.tables

    def write(self, table_id, data, start_pos):
        _table = self._tables[table_id]
        r_num, c_num = data.shape
        r1st, c1st = start_pos
        rows = _table.rows[r1st:r1st+r_num]
        cols_num = len(rows[0].cells)
        data_row_cursor = 0
        for row in rows:
            left = 0
            right = 1
            count = 0
            data_column_cursor = 0
            cells = row.cells
            while right < cols_num and count < c1st + c_num:
                if id(cells[left]) == id(cells[right]):
                    right += 1
                    continue
                else:
                    left = right
                    right += 1
                    count += 1
                    if count < c1st:
                        continue
                    else:
                        cells[left].text = data[data_row_cursor, data_column_cursor]
                        data_column_cursor += 1
            data_row_cursor += 1

    def test_save(self):
        import time
        time_stamp = time.strftime('%m%d%H%M%S', time.localtime())
        self._doc.save('test/test'+ time_stamp +'.docx')

    def save(self):
        pass



if __name__ == '__main__':
    writer = Writer('data/1抽检原始记录A1级三相外置.docx')
    # for i in range(10):
        # writer.write(3, None, (4+i//4,6+i%4))
    import numpy as np
    data = np.arange(15*4).reshape(15,4).astype(str)
    writer.write(3, data, (4,3))
    writer.test_save()