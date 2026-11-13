import unittest

from pydantic import ValidationError
from main import DiplomeRequest


class TestDatesNaissance(unittest.TestCase):
    def payload(self, date):
        return dict(niveau="DT3", matricule="TEST-1", nom="Candidat Test", mention="Bien",
                    date_naissance=date, lieu_naissance="Antananarivo")

    def test_dates_impossibles_refusees(self):
        for date in ("31/02/1990", "29/02/1991", "31/04/1990", "00/01/1990", "01/13/1990", "01/01/0000", "vers 0000"):
            with self.subTest(date=date), self.assertRaises(ValidationError):
                DiplomeRequest(**self.payload(date))

    def test_date_bissextile_et_date_approximative(self):
        for value, attendu in (("29/02/1992", "29/02/1992"), ("1990", "vers 1990"),
                               ("vers l'année 1990", "vers 1990")):
            self.assertEqual(DiplomeRequest(**self.payload(value)).date_naissance, attendu)


if __name__ == "__main__":
    unittest.main()
