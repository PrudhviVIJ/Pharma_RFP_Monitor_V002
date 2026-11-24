import json
import os
import unittest
from unittest.mock import patch
from tempfile import TemporaryDirectory

from gemini_enrichment import (
    _parse_json,
    extract_candidate_links,
    merge_enrichment,
    resolve_application_url,
    DiscoveredAnnouncement,
    save_discovered,
)
from quality_gate import detect_login_only_page, verify_opportunity
from main import (
    build_email_alert_rows,
    classify,
    deadline_is_expired,
    deadline_is_too_soon,
    extract_deadline_text,
    is_target_rfp_record,
    send_email_alert,
)
from report_csv import read_records, write_records
from opportunity_report import render_markdown


class GeminiEnrichmentTests(unittest.TestCase):
    def test_classification_requires_target_topic_and_rfp_language(self):
        oncology = classify(
            "Example Pharma", "https://example.com/oncology-rfp", "",
            "Request for Proposals for oncology research. Applications are open until January 1, 2027.",
        )
        generic = classify(
            "Example Pharma", "https://example.com/grants", "",
            "Funding opportunities are available for qualified organizations.",
        )
        neurology_grant = classify(
            "Example Pharma", "https://example.com/neurology", "",
            "We provide grants for neurology education. Applications are open.",
        )
        self.assertTrue(oncology["is_target_rfp"])
        self.assertTrue(oncology["flag"])
        self.assertFalse(generic["flag"])
        self.assertFalse(neurology_grant["flag"])

    def test_target_rfp_record_requires_both_scope_conditions(self):
        self.assertTrue(is_target_rfp_record({
            "opportunity_title": "Neurology RFP",
            "evidence": "Request for proposals for neurology education.",
        }))
        self.assertFalse(is_target_rfp_record({
            "opportunity_title": "General RFP",
            "evidence": "Request for proposals for diabetes education.",
        }))

    def test_parse_json_code_fence(self):
        self.assertEqual(_parse_json('```json\n{"announcements": []}\n```'), {"announcements": []})

    def test_resolves_direct_rfp_pdf(self):
        html = '''
        <a href="/about">About</a>
        <a href="/download/JJIM_Oncology_Nursing_Fellowship_RFP.pdf">
          Download Oncology Nursing Fellowship RFP
        </a>
        '''
        result = resolve_application_url("https://www.jnj.com/grants/educational", html)
        self.assertEqual(
            result,
            "https://www.jnj.com/download/JJIM_Oncology_Nursing_Fellowship_RFP.pdf",
        )

    def test_saves_discovery_records(self):
        from tempfile import TemporaryDirectory
        with TemporaryDirectory() as directory:
            path = f"{directory}/discovered.json"
            save_discovered([DiscoveredAnnouncement("Example", "https://example.com")], path)
            with open(path, encoding="utf-8") as handle:
                self.assertEqual(json.load(handle)[0]["company"], "Example")

    def test_extracts_bounded_same_site_candidate_links(self):
        html = '''
        <a href="/about">About</a>
        <a href="/grants/oncology">Oncology grant opportunities</a>
        <a href="https://other.example/rfp.pdf">Other site</a>
        <a href="/files/current-rfp.pdf">Download RFP</a>
        '''
        links = extract_candidate_links("https://example.com/programs", html, max_links=5)
        self.assertEqual(links[0], "https://example.com/files/current-rfp.pdf")
        self.assertIn("https://example.com/grants/oncology", links)
        self.assertNotIn("https://other.example/rfp.pdf", links)

    def test_merges_gemini_details_without_erasing_scraped_data(self):
        result = {"page_title": "Landing page", "evidence": "RFP announced", "funding_terms": "rfp", "flag": False,
                  "application_url": "https://example.com/scraped-rfp.pdf"}
        merge_enrichment(result, {"relevant": True, "confidence": "high", "opportunity_title": "Oncology RFP",
                                  "deadline_text": "25 Nov 2026", "application_url": "https://example.com/rfp.pdf",
                                  "funding_terms": ["CGA"], "evidence": "Applications are open", "summary": "Open oncology call"})
        self.assertTrue(result["flag"])
        self.assertEqual(result["application_url"], "https://example.com/scraped-rfp.pdf")
        self.assertIn("rfp", result["funding_terms"])
        self.assertIn("CGA", result["funding_terms"])
        self.assertEqual(result["gemini_confidence"], "high")

    def test_rejects_unrelated_policy_document(self):
        result = verify_opportunity(
            "https://example.com/grants",
            "https://example.com/cookie-policy.pdf",
            "Cookie Policy",
            "Cookie Policy. We use cookies to improve the website. Privacy settings.",
        )
        self.assertEqual(result["document_status"], "UNRELATED_DOCUMENT")
        self.assertEqual(result["email_eligible"], "No")

    def test_accepts_proposal_with_actionable_details(self):
        result = verify_opportunity(
            "https://example.com/grants",
            "https://example.com/oncology-rfp.pdf",
            "2027 Oncology RFP",
            "Request for Proposals for oncology research. Applications are open. "
            "Eligible institutions must submit a proposal by January 1, 2027. "
            "Funding awards are available up to $100,000. Please upload the proposal.",
        )
        self.assertEqual(result["document_status"], "VERIFIED_PROPOSAL")
        self.assertEqual(result["email_eligible"], "Yes")
        self.assertTrue(result["evidence_quote"])

    def test_does_not_accept_incidental_term_mention(self):
        result = verify_opportunity(
            "https://example.com/about",
            "https://example.com/about",
            "About us",
            "Our website footer mentions independent medical education and external funding. "
            "This page contains no application, deadline, eligibility, or submission instructions.",
        )
        self.assertNotEqual(result["document_status"], "VERIFIED_PROPOSAL")
        self.assertEqual(result["email_eligible"], "No")

    def test_rejects_stale_cures_within_reach_rfp(self):
        result = verify_opportunity(
            "https://www.cureswithinreach.org/wp-content/uploads/2023/05/CWR-ReGRoW-RFP-2023.pdf",
            "https://www.cureswithinreach.org/wp-content/uploads/2023/05/CWR-ReGRoW-RFP-2023.pdf",
            "2023 ReGRoW RFP",
            "This Request for Proposals is accepting clinical repurposing trials. "
            "The LOI submission deadline is July 12, 2023. Eligible institutions must submit a proposal.",
        )
        self.assertEqual(result["document_status"], "STALE_DOCUMENT")
        self.assertEqual(result["email_eligible"], "No")

    def test_rejects_policy_and_compliance_documents(self):
        bms = verify_opportunity(
            "https://www.bms.com/assets/UK-CA-Slavery-Human-Trafficking-Statement.pdf",
            "https://www.bms.com/assets/UK-CA-Slavery-Human-Trafficking-Statement.pdf",
            "Modern Slavery & Human Trafficking Statement",
            "This statement addresses modern slavery, forced labor, and human trafficking across the supply chain.",
        )
        exact = verify_opportunity(
            "https://www.exactsciences.com/legal/ca-compliance-program.pdf",
            "https://www.exactsciences.com/legal/ca-compliance-program.pdf",
            "California Compliance Declaration",
            "The company maintains a Health Care Compliance Program and written compliance policies.",
        )
        self.assertEqual(bms["email_eligible"], "No")
        self.assertEqual(exact["email_eligible"], "No")

    def test_rejects_old_proposal_review_summary(self):
        result = verify_opportunity(
            "https://cdn.pfizer.com/rfp_files/Full_Proposal_Review_Summary_Smoking_Cessation.pdf",
            "https://cdn.pfizer.com/rfp_files/Full_Proposal_Review_Summary_Smoking_Cessation.pdf",
            "Full Proposal Review Summary - Smoking Cessation",
            "The External Review Panel met on September 12, 2012. Pfizer Independent Grants for Learning and Change Request for Proposals: Smoking Cessation.",
        )
        self.assertEqual(result["email_eligible"], "No")

    def test_rejects_login_only_portal(self):
        result = verify_opportunity(
            "https://www.cybergrants.com/pls/cybergrants/ao_login.login?x_gm_id=9892",
            "https://www.cybergrants.com/pls/cybergrants/ao_login.login?x_gm_id=9892",
            "Independent Medical Education Contributions",
            "Requests for Proposals may be posted throughout the year. Apply now through the portal.",
        )
        self.assertEqual(result["email_eligible"], "No")

    def test_detects_login_page_for_grounded_alternative_search(self):
        self.assertEqual(
            detect_login_only_page(
                "https://www.cybergrants.com/pls/cybergrants/ao_login.login",
                "CyberGrants Login",
                "Enter your username and password. Forgot password?",
            ),
            "login-only portal",
        )

    def test_unconfigured_email_is_optional(self):
        with patch.dict(os.environ, {
            "SMTP_HOST": "",
            "ALERT_FROM": "",
            "ALERT_TO": "",
        }, clear=False):
            self.assertIsNone(send_email_alert([{"company": "Example"}]))

    def test_email_rows_split_flattened_table_evidence_and_use_specialty(self):
        rows = build_email_alert_rows([{
            "company": "BMS",
            "source_url": "https://example.com/rfp",
            "final_url": "https://example.com/rfp",
            "application_url": "",
            "opportunity_type": "RFP",
            "opportunity_title": "Request for education",
            "email_eligible": "Yes",
            "email_block_reason": "",
            "access_issue": "",
            "deadline_text": "January 18, 2027",
            "deadline_date": "2027-01-18",
            "evidence_quote": (
                "RFP Code RFE-27-IM-101; Area of Interest Pulmonary Fibrosis; "
                "LOR Due Date January 18, 2027"
            ),
            "new_evidence": [],
            "therapeutic_areas": "",
            "focus": "",
            "topic_terms": [],
        }])
        self.assertEqual(len(rows), 1)
        self.assertIn("Pulmonary Fibrosis", rows[0][2])
        self.assertIn("pulmonary fibrosis", rows[0][5].casefold())
        self.assertNotIn("**", rows[0][2])

    def test_deadlines_older_than_yesterday_are_expired(self):
        from datetime import date

        today = date(2026, 11, 19)
        self.assertTrue(deadline_is_expired({"deadline_text": "November 17, 2026"}, today))
        self.assertTrue(deadline_is_expired({"deadline_text": "November 18, 2026"}, today))
        self.assertFalse(deadline_is_expired({"deadline_text": "November 19, 2026"}, today))
        self.assertFalse(deadline_is_expired({"deadline_text": "November 20, 2026"}, today))
        self.assertFalse(deadline_is_expired({"deadline_text": "No deadline stated"}, today))

    def test_normalizes_malformed_table_dates(self):
        text = "LOR Due Date January 18,; 2026 - 5PM EST"
        extracted = extract_deadline_text(text)
        self.assertIn("January 18, 2026", extracted)

    def test_blocks_deadlines_within_minimum_lead_time(self):
        from datetime import date

        today = date(2026, 11, 24)
        self.assertTrue(deadline_is_too_soon({"deadline_text": "November 24, 2026"}, today))
        self.assertTrue(deadline_is_too_soon({"deadline_text": "December 1, 2026"}, today))
        self.assertFalse(deadline_is_too_soon({"deadline_text": "December 2, 2026"}, today))

    def test_report_write_is_atomic_and_readable(self):
        with TemporaryDirectory() as directory:
            path = f"{directory}/report.csv"
            write_records(path, [{"company": "Example", "flag": True}], ["company", "flag"])
            self.assertEqual(read_records(path), [{"company": "Example", "flag": "True"}])

    def test_markdown_report_groups_verified_opportunities(self):
        report = render_markdown([{
            "company": "Example Pharma",
            "sponsor": "Example Pharma",
            "opportunity_title": "Neurology RFP",
            "opportunity_status": "OPEN",
            "opportunity_type": "RFP",
            "therapeutic_areas": "neurology",
            "deadline_text": "December 1, 2026",
            "source_url": "https://example.com/grants",
            "final_url": "https://example.com/grants",
            "document_status": "VERIFIED_PROPOSAL",
            "flag": True,
            "priority": "HIGH",
            "evidence_quote": "Applications are open until December 1, 2026.",
        }], "November 11, 2026")
        self.assertIn("# RFP and funding opportunities found", report)
        self.assertIn("## Neurology opportunities", report)
        self.assertIn("Example Pharma – Neurology RFP", report)


if __name__ == "__main__":
    unittest.main()
