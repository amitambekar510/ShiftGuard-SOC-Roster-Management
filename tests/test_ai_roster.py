import unittest
from ai_roster import create

def example():
    return {"year":2026,"month":10,"minimum_per_shift":{"SOC L1":1,"SOC L2":1},"analysts":[{"name":"Lead","role":"Team Lead","priority":"general"}]+[{"name":f"L1-{i}","role":"SOC L1","priority":"soc"} for i in range(5)]+[{"name":f"L2-{i}","role":"SOC L2","priority":"soc"} for i in range(5)]}

class Tests(unittest.TestCase):
    def test_general_priority(self):
        roster=create(example())[2]
        self.assertEqual(roster["Lead"]["2026-10-03"],"Week Off")
        self.assertEqual(roster["Lead"]["2026-10-05"],"General")
    def test_shortage_reported(self):
        c=example()
        c["analysts"]=[p for p in c["analysts"] if p["name"] in ("Lead","L1-0","L2-0")]
        self.assertTrue(create(c)[3])
    def test_duplicate_rejected(self):
        c=example()
        c["analysts"].append(c["analysts"][0])
        with self.assertRaises(ValueError):create(c)
