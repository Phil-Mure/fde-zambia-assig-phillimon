from django.test import Client, TestCase

from bulletin.services import build_bulletin


class BulletinServiceTests(TestCase):
    def test_build_bulletin_returns_expected_sections(self):
        report = build_bulletin("2025Q1")
        self.assertEqual(report.current_period, "2025Q1")
        self.assertEqual(len(report.top_facilities), 10)
        self.assertIn("deliveries", report.maternal_summary)
        self.assertEqual(len(report.performance_scores), 15)
        self.assertGreater(len(report.trend_analysis), 0)

    def test_top_facility_is_highest_volume(self):
        report = build_bulletin("2025Q1")
        self.assertEqual(report.top_facilities[0]["facility_name"], "Kigali University Teaching Hospital")


class BulletinViewTests(TestCase):
    def setUp(self):
        self.client = Client()

    def test_dashboard_returns_200(self):
        response = self.client.get("/")
        self.assertEqual(response.status_code, 200)
        self.assertContains(response, "Quarterly Health Bulletin")

    def test_json_api_returns_metrics(self):
        response = self.client.get("/api/bulletin/?period=2025Q1")
        self.assertEqual(response.status_code, 200)
        payload = response.json()
        self.assertEqual(payload["current_period"], "2025Q1")
        self.assertIn("maternal_summary", payload)
