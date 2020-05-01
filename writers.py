from docx import Document
import time


class Writer:
    def __init__(self, file_path):
        self._doc = Document(file_path)
        self._tables = self._doc.tables

    def write(self, table_id, data, start_pos):
        _table = self._tables[table_id]
        # _table.cell(start_pos[0], start_pos[1]).text = 'hello world'
        # self._doc.save('data/test'+time.strftime('%m%d%H%M%S', time.localtime())+'.docx')
        r_num, c_num = data.shape
        r1st, c1st = start_pos
        rows = _table.rows[r1st:r1st+r_num]
        for row in rows:
            # cursor1 = 0
            # cursor2 = 1
            # while count < c1st:
            #     if row.cells[cursor1] == row.cells[cursor2]:
            #         cursor2 += 1
            #     else:
            #         cursor1 = cursor2
            #         count += 1
            #         cursor2 += 1

            row_cells_temp = row.cells
            print(len(row_cells_temp))
            row_cells = list(set(row_cells_temp))
            row_cells.sort(key=row_cells_temp.index)
            for c in row_cells[c1st:c1st+c_num]:
                pass


if __name__ == '__main__':
    writer = Writer('data/1抽检原始记录A1级三相外置.docx')
    # for i in range(10):
        # writer.write(3, None, (4+i//4,6+i%4))
    import numpy as np
    data = np.zeros((5,4))
    writer.write(3, data, (4,3))