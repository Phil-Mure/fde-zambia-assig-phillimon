from unittest.mock import patch

from django.test import Client, TestCase
from django.template.loader import render_to_string

from bulletin.services import AUTO_REFRESH_SECONDS, available_periods, build_bulletin, load_datasets


class BulletinServiceTests(TestCase):
    def test_build_bulletin_returns_expected_sections(self):
        latest_period = available_periods()[-1].key
        report = build_bulletin(latest_period)
        self.assertEqual(report.current_period, latest_period)
        self.assertGreaterEqual(len(report.source_files), 5)
        self.assertGreaterEqual(report.source_summary["files_loaded"], 5)
        self.assertIn("synced_at", report.source_summary)
        self.assertIn("file_catalog", report.source_summary)
        self.assertEqual(len(report.top_facilities), 10)
        self.assertIn("neonatal_mortality_rate", report.maternal_summary)
        self.assertGreaterEqual(len(report.performance_scores), 10)
        self.assertGreater(len(report.provincial_overview), 0)
        self.assertGreater(len(report.risk_alerts), 0)
        self.assertGreater(len(report.trend_analysis), 0)

    def test_file_catalog_classifies_core_drive_files(self):
        datasets = load_datasets()
        roles = {item["name"]: item["role"] for item in datasets["file_catalog"]}
        self.assertEqual(roles["clinical_neonatal.csv"], "clinical")
        self.assertEqual(roles["facilities.csv"], "facilities")
        self.assertEqual(roles["governance.csv"], "governance")
        self.assertEqual(roles["healthcare_workers.csv"], "healthcare_workers")
        self.assertEqual(roles["operations.csv"], "operations")

    def test_period_options_are_discovered_from_clinical_data(self):
        periods = available_periods()
        self.assertGreaterEqual(len(periods), 3)
        self.assertTrue(periods[-1].key.startswith("2024Q"))
        self.assertIn("2024Q", periods[0].label)

    def test_monthly_quarterly_and_annual_reports_are_supported(self):
        monthly_key = available_periods(granularity="M")[-1].key
        monthly = build_bulletin(monthly_key, granularity="M")
        quarterly = build_bulletin("2024Q4", granularity="Q")
        annual = build_bulletin("2024", granularity="A")

        self.assertEqual(monthly.granularity, "M")
        self.assertEqual(quarterly.granularity, "Q")
        self.assertEqual(annual.granularity, "A")
        self.assertEqual(len(quarterly.period_breakdown), quarterly.rollup_summary["selected_months_loaded"])
        self.assertLess(quarterly.rollup_summary["selected_reporting_completeness_pct"], 100)
        self.assertEqual(len(annual.period_breakdown), 4)
        self.assertEqual(
            annual.rollup_summary["quarter_sum_deliveries"],
            annual.maternal_summary["deliveries"],
        )
        self.assertGreater(len(annual.rollup_summary["monthly_contributions"]), 0)
        self.assertGreater(len(annual.advanced_aggregations), 8)

    def test_top_facility_is_ranked_by_current_delivery_volume(self):
        report = build_bulletin("2024Q4")
        self.assertEqual(report.top_facilities[0]["facility_name"], "Kicukiro Health Center 2")
        self.assertGreater(report.top_facilities[0]["deliveries"], report.top_facilities[1]["deliveries"])


class BulletinViewTests(TestCase):
    def setUp(self):
        self.client = Client()

    def test_dashboard_returns_200(self):
        response = self.client.get("/")
        self.assertEqual(response.status_code, 200)
        self.assertContains(response, "Rwanda Neonatal Intelligence Dashboard")
        self.assertContains(response, "Serving analytics from local source files")
        self.assertContains(response, 'rel="icon"')
        self.assertContains(response, 'href="/static/bulletin/favicon.svg"')

    def test_force_refresh_query_param(self):
        response = self.client.get("/api/bulletin/?period=2024Q4&refresh=1")
        self.assertEqual(response.status_code, 200)
        payload = response.json()
        self.assertIn("synced_at", payload["source_summary"])
        self.assertIn("sync_status", payload["source_summary"])
        self.assertNotIn("FileURLRetrievalError", payload["source_summary"]["sync_message"])
        self.assertNotIn("temporarily unavailable", payload["source_summary"]["sync_message"])

    def test_500_template_has_support_contact_and_favicon(self):
        html = render_to_string("500.html")
        self.assertIn("phillimon@sandtech.com", html)
        self.assertIn("bulletin/favicon.svg", html)
        self.assertIn("Something went wrong", html)

    def test_favicon_route_redirects_to_static_svg(self):
        response = self.client.get("/favicon.ico")
        self.assertEqual(response.status_code, 301)
        self.assertEqual(response["Location"], "/static/bulletin/favicon.svg")


class BulletinSyncFallbackTests(TestCase):
    def test_drive_sync_failure_falls_back_to_local_cache(self):
        with patch("bulletin.services._download_drive_folder", side_effect=RuntimeError("rate limited")):
            report = build_bulletin("2024Q4", force_refresh=True)
        self.assertGreaterEqual(report.source_summary["files_loaded"], 5)
        self.assertEqual(report.source_summary["sync_status"], "synced")
        self.assertIn("local", report.source_summary["sync_message"].lower())

    def test_json_api_returns_metrics(self):
        response = self.client.get("/api/bulletin/?granularity=A&period=2024")
        self.assertEqual(response.status_code, 200)
        payload = response.json()
        self.assertEqual(payload["current_period"], "2024")
        self.assertEqual(payload["granularity"], "A")
        self.assertIn("maternal_summary", payload)
        self.assertIn("rollup_summary", payload)
        self.assertIn("advanced_aggregations", payload)
        self.assertIn("source_files", payload)
