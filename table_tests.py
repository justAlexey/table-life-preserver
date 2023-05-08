import unittest
import table

import pandas

class TableTest(unittest.TestCase):

    def test_get_file_list(self):
        cm_list = ['CM134747.xlsm', 'CM135654.1.xlsm', 'CM135654.2.xlsm', 'CM135654.3.xlsm', 'CM135655.1.xlsm','CM135655.2.xlsm','CM135655.3.xlsm','CM135724.xlsm']
        self.assertEqual(table.get_file_list("./archive"),cm_list)
        self.assertEqual(table.get_file_list("./not_a_path"),-1)

    def test_get_serial_numbers_by_file(self):
        numbers_list = ['L2988976', 'L547108', 'L640588', 'L2024971', 'L2988950', 'L2988796', 'L639718', 'L547080', 'L1807567', 'L2988798', 'L993066', 'L2988957', 'L4926549', 'Lnan', 'Lnan', 'Lnan', 'Lnan', 'Lnan']
        self.assertEqual(table.get_serial_numbers_by_file('./archive/','CM134747.xlsm'),numbers_list)



def get_serial_numbers_by_file(path:str, file: str)->list:
    """Возвращает список серийных номеров из файла"""
    print(f"Извлекаю номера из {file}")
    temp = pandas.read_excel(f"{path}{file}")
    temp = temp.iloc[6:25, [3]].squeeze().tolist()
    print(type(temp[14]))
    return ["L"+str(x) for x in temp]


if __name__=="__main__":
    print(get_serial_numbers_by_file('./archive/','CM134747.xlsm'))
    # unittest.main()