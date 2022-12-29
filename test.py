import unittest
from vacancies import Salary, Vacancy, InputConect
from prettytable import PrettyTable


class Test(unittest.TestCase):
    def setUp(self):
        self.vacancy = Vacancy('Аналитик', 'Первый', ['HTML', 'PHP', 'CSS', 'JavaScript', 'SQL'],
                                'between3And6', 'True', 'hh', Salary('40000', '60000', 'False', 'RUR'),
                                'Владивосток', '2022-07-06T02:05:26+0300')
        self.table = PrettyTable()
        self.fields1 = ''
        self.fields2 = 'Название, Описание'

    def test_formatter_date(self):
        dict_vacancies = [self.vacancy]
        self.assertEqual(InputConect.formatter(dict_vacancies)[0]['Дата публикации вакансии'], '06.07.2022')

    def test_formatter_skills(self):
        dict_vacancies = [self.vacancy]
        self.assertEqual(InputConect.formatter(dict_vacancies)[0]['Навыки'], 'HTML\nPHP\nCSS\nJavaScript\nSQL')

    def test_formatter_experience_id(self):
        dict_vacancies = [self.vacancy]
        self.assertEqual(InputConect.formatter(dict_vacancies)[0]['Опыт работы'], 'От 3 до 6 лет')

    def test_formatter_true_false(self):
        dict_vacancies = [self.vacancy]
        self.assertEqual(InputConect.formatter(dict_vacancies)[0]['Премиум-вакансия'], 'Да')

    def test_formatter_text(self):
        dict_vacancies = [self.vacancy]
        self.assertEqual(InputConect.formatter(dict_vacancies)[0]['Название'], 'Аналитик')

    def test_fields_table_none(self):
        self.table.field_names = ['№', 'Название', 'Навыки']
        self.assertEqual(InputConect.get_fields_table(self.table, self.fields1), ['№', 'Название', 'Навыки'])

    def test_fields_table_(self):
        self.table.field_names = ['№', 'Название', 'Навыки']
        self.assertEqual(InputConect.get_fields_table(self.table, self.fields2), ['№', 'Название', 'Описание'])

if __name__ == '__main__':
    unittest.main()