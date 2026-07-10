import json

from django.http import HttpRequest, HttpResponse, JsonResponse
from django.shortcuts import render
from django.views.decorators.http import require_GET

from .services import AUTO_REFRESH_SECONDS, build_bulletin, report_to_dict


def _report_context(request: HttpRequest, period: str | None = None) -> dict[str, object]:
    force_refresh = request.GET.get("refresh") == "1"
    granularity = request.GET.get("granularity")
    report = build_bulletin(period, granularity=granularity, force_refresh=force_refresh)
    return {
        "report": report,
        "available_periods": report.available_periods,
        "available_granularities": report.available_granularities,
        "auto_refresh_seconds": AUTO_REFRESH_SECONDS,
        "force_refreshed": force_refresh,
    }


@require_GET
def dashboard(request: HttpRequest) -> HttpResponse:
    """HTML bulletin dashboard for MoH stakeholders."""
    period = request.GET.get("period")
    context = _report_context(request, period)
    return render(request, "bulletin/dashboard.html", context)


@require_GET
def bulletin_json(request: HttpRequest) -> JsonResponse:
    """Machine-readable output for downstream Analytics Template Toolkit / Superset."""
    period = request.GET.get("period")
    context = _report_context(request, period)
    report = context["report"]
    payload = report_to_dict(report)
    return JsonResponse(payload, json_dumps_params={"indent": 2})


@require_GET
def bulletin_export(request: HttpRequest) -> HttpResponse:
    """Download bulletin as JSON file (stand-in for PDF/Excel export in production)."""
    period = request.GET.get("period")
    context = _report_context(request, period)
    report = context["report"]
    payload = report_to_dict(report)
    response = HttpResponse(
        json.dumps(payload, indent=2),
        content_type="application/json",
    )
    response["Content-Disposition"] = (
        f'attachment; filename="quarterly_health_bulletin_{report.current_period}.json"'
    )
    return response
