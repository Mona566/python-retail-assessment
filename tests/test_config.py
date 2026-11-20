"""Tests for environment-backed config helpers."""

from __future__ import annotations

from src.config import candidate_name, poc_test_email, poc_test_emails, submission_email


def test_candidate_name_env_branches(monkeypatch):
    monkeypatch.setenv("CANDIDATE_NAME", "Ada Lovelace")
    assert candidate_name() == "Ada Lovelace"
    monkeypatch.setenv("CANDIDATE_NAME", "   ")
    assert candidate_name() == "Candidate"
    monkeypatch.delenv("CANDIDATE_NAME", raising=False)
    assert candidate_name() == "Candidate"


def test_submission_email_env_branches(monkeypatch):
    monkeypatch.setenv("SUBMISSION_EMAIL", "hire@example.com")
    assert submission_email() == "hire@example.com"
    monkeypatch.setenv("SUBMISSION_EMAIL", "  ")
    assert submission_email() is None
    monkeypatch.delenv("SUBMISSION_EMAIL", raising=False)
    assert submission_email() is None


def test_poc_test_emails_parsing(monkeypatch):
    monkeypatch.delenv("POC_TEST_EMAIL", raising=False)
    assert poc_test_emails() == []
    assert poc_test_email() is None

    monkeypatch.setenv("POC_TEST_EMAIL", "a@x.com, b@y.com;c@z.com")
    assert poc_test_emails() == ["a@x.com", "b@y.com", "c@z.com"]
    assert poc_test_email() == "a@x.com"

    monkeypatch.setenv("POC_TEST_EMAIL", "  , ; ")
    assert poc_test_emails() == []
