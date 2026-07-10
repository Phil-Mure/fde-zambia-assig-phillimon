import json

from django.http import HttpRequest, HttpResponse, JsonResponse
from django.shortcuts import render
from django.views.decorators.http import require_GET

from .services import build_bulletin


@require_GET
def dashboard(request: HttpRequest) -> HttpResponse:
    """HTML bulletin dashboard for MoH stakeholders."""
    period = request.GET.get("period")
    report = build_bulletin(period)
    return render(
        request,
        "bulletin/dashboard.html",
        {
            "report": report,
            "available_periods": ["2024Q4", "2025Q1"],
        },
    )


@require_GET
def bulletin_json(request: HttpRequest) -> JsonResponse:
    """Machine-readable output for downstream Analytics Template Toolkit / Superset."""
    period = request.GET.get("period")
    report = build_bulletin(period)
    payload = {
        "current_period": report.current_period,
        "previous_period": report.previous_period,
        "generated_at": report.generated_at,
        "top_facilities": report.top_facilities,
        "maternal_summary": report.maternal_summary,
        "performance_scores": report.performance_scores,
        "trend_analysis": report.trend_analysis,
    }
    return JsonResponse(payload, json_dumps_params={"indent": 2})


@require_GET
def bulletin_export(request: HttpRequest) -> HttpResponse:
    """Download bulletin as JSON file (stand-in for PDF/Excel export in production)."""
    period = request.GET.get("period")
    report = build_bulletin(period)
    payload = {
        "current_period": report.current_period,
        "previous_period": report.previous_period,
        "generated_at": report.generated_at,
        "top_facilities": report.top_facilities,
        "maternal_summary": report.maternal_summary,
        "performance_scores": report.performance_scores,
        "trend_analysis": report.trend_analysis,
    }
    response = HttpResponse(
        json.dumps(payload, indent=2),
        content_type="application/json",
    )
    response["Content-Disposition"] = (
        f'attachment; filename="quarterly_health_bulletin_{report.current_period}.json"'
    )
    return response
