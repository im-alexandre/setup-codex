#!/usr/bin/env python
"""Provision and deploy a GitHub-backed Docker Compose app on Coolify."""

from __future__ import annotations

import argparse
import base64
import json
import os
import secrets
import sys
import time
from typing import Any

try:
    import requests
except ImportError as exc:  # pragma: no cover - runtime guidance
    raise SystemExit("Missing dependency: install requests in the active Python environment.") from exc


TERMINAL_STATUSES = ("finished", "success", "failed", "cancelled", "error")


def token_secret(num_bytes: int) -> str:
    raw = secrets.token_bytes(num_bytes)
    return base64.urlsafe_b64encode(raw).decode("ascii").rstrip("=")


def parse_env(values: list[str]) -> dict[str, str]:
    parsed: dict[str, str] = {}
    for item in values:
        if "=" not in item:
            raise SystemExit(f"Invalid --env value {item!r}; expected KEY=VALUE.")
        key, value = item.split("=", 1)
        key = key.strip()
        if not key:
            raise SystemExit(f"Invalid --env value {item!r}; empty key.")
        parsed[key] = value
    return parsed


class Coolify:
    def __init__(self, base_url: str, api_key: str) -> None:
        self.base_url = base_url.rstrip("/")
        self.session = requests.Session()
        self.session.headers.update(
            {
                "Authorization": f"Bearer {api_key}",
                "Accept": "application/json",
            }
        )

    def request(self, method: str, path: str, *, json_body: Any | None = None) -> Any:
        url = f"{self.base_url}{path}"
        response = self.session.request(method, url, json=json_body, timeout=90)
        try:
            response.raise_for_status()
        except requests.HTTPError as exc:
            detail = response.text[:1000]
            raise SystemExit(f"Coolify API {method} {path} failed: {response.status_code} {detail}") from exc
        if not response.content:
            return None
        return response.json()

    def get(self, path: str) -> Any:
        return self.request("GET", path)

    def post(self, path: str, body: Any) -> Any:
        return self.request("POST", path, json_body=body)

    def patch(self, path: str, body: Any) -> Any:
        return self.request("PATCH", path, json_body=body)

    def delete(self, path: str) -> Any:
        return self.request("DELETE", path)


def find_first(items: list[Any], predicate) -> Any | None:
    for item in items:
        if predicate(item):
            return item
    return None


def env_payload(envs: dict[str, str]) -> dict[str, Any]:
    return {
        "data": [
            {
                "key": key,
                "value": value,
                "is_preview": False,
                "is_literal": False,
                "is_multiline": False,
                "is_shown_once": key in {"POSTGRES_PASSWORD", "SECRET_KEY"},
            }
            for key, value in envs.items()
        ]
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--coolify-url", default=os.environ.get("COOLIFY_URL", "https://coolify.drg.ink"))
    parser.add_argument("--project-name", required=True)
    parser.add_argument("--environment-name", required=True)
    parser.add_argument("--repository", required=True)
    parser.add_argument("--branch", required=True)
    parser.add_argument("--domain", required=True)
    parser.add_argument("--service-name", required=True)
    parser.add_argument("--server-uuid", required=True)
    parser.add_argument("--github-app-uuid", required=True)
    parser.add_argument("--compose-location", default="/docker-compose.yml")
    parser.add_argument("--api-path", default="")
    parser.add_argument("--ports-exposes", default="80")
    parser.add_argument("--env", action="append", default=[], help="Runtime env var as KEY=VALUE. Repeatable.")
    parser.add_argument("--django-default-envs", action="store_true")
    parser.add_argument("--skip-deploy", action="store_true")
    parser.add_argument("--keep-auto-production-environment", action="store_true")
    args = parser.parse_args()

    api_key = os.environ.get("COOLIFY_API_KEY")
    if not api_key:
        raise SystemExit("COOLIFY_API_KEY is required in the environment.")

    client = Coolify(args.coolify_url, api_key)
    domain_host = requests.utils.urlparse(args.domain).hostname or args.domain.replace("https://", "").replace("http://", "")

    version = client.get("/api/v1/version")
    print(f"coolify_version={version}")

    projects = client.get("/api/v1/projects")
    project = find_first(projects, lambda item: item.get("name") == args.project_name)
    if project is None:
        project = client.post(
            "/api/v1/projects",
            {"name": args.project_name, "description": "Created by Codex coolify-deploy skill"},
        )
        print(f"created_project_uuid={project['uuid']}")
    else:
        print(f"existing_project_uuid={project['uuid']}")
    project_uuid = project["uuid"]

    environments = client.get(f"/api/v1/projects/{project_uuid}/environments")
    environment = find_first(environments, lambda item: item.get("name") == args.environment_name)
    if environment is None:
        environment = client.post(
            f"/api/v1/projects/{project_uuid}/environments",
            {"name": args.environment_name},
        )
        print(f"created_environment={args.environment_name} uuid={environment['uuid']}")
    else:
        print(f"existing_environment={args.environment_name} uuid={environment['uuid']}")

    if not args.keep_auto_production_environment and args.environment_name != "production":
        environments = client.get(f"/api/v1/projects/{project_uuid}/environments")
        auto_prod = find_first(environments, lambda item: item.get("name") == "production")
        if auto_prod is not None:
            try:
                client.delete(f"/api/v1/projects/{project_uuid}/environments/{auto_prod['uuid']}")
                print("deleted_empty_environment=production")
            except SystemExit:
                print("kept_environment=production reason=delete_failed_or_not_empty")

    apps = client.get("/api/v1/applications")
    app = find_first(
        apps,
        lambda item: item.get("git_repository") == args.repository
        or domain_host in str(item.get("docker_compose_domains", ""))
        or args.project_name.lower() in str(item.get("name", "")).lower(),
    )
    if app is None:
        app = client.post(
            "/api/v1/applications/private-github-app",
            {
                "project_uuid": project_uuid,
                "server_uuid": args.server_uuid,
                "environment_name": args.environment_name,
                "github_app_uuid": args.github_app_uuid,
                "git_repository": args.repository,
                "git_branch": args.branch,
                "build_pack": "dockercompose",
                "docker_compose_location": args.compose_location,
                "ports_exposes": args.ports_exposes,
                "docker_compose_domains": [{"name": args.service_name, "domain": args.domain}],
                "is_auto_deploy_enabled": True,
                "is_force_https_enabled": True,
                "instant_deploy": False,
            },
        )
        print(f"created_app_uuid={app['uuid']}")
    else:
        print(f"existing_app_uuid={app['uuid']}")
    app_uuid = app["uuid"]

    runtime_env = parse_env(args.env)
    if args.django_default_envs:
        runtime_env = {
            "SQL_ENGINE": "django.db.backends.postgresql",
            "SQL_HOST": "db_electre",
            "SQL_PORT": "5432",
            "POSTGRES_USER": "electremor",
            "POSTGRES_PASSWORD": token_secret(36),
            "POSTGRES_DB": "electremor",
            "SECRET_KEY": token_secret(48),
            "DEBUG": "1",
            "DJANGO_ALLOWED_HOSTS": f"localhost,127.0.0.1,coolify.drg.ink,{domain_host}",
            **runtime_env,
        }
    if runtime_env:
        updated = client.patch(f"/api/v1/applications/{app_uuid}/envs/bulk", env_payload(runtime_env))
        keys = ",".join(sorted({item["key"] for item in updated}))
        print(f"updated_env_keys={keys}")

    deployment_uuid = None
    deployment = None
    if not args.skip_deploy:
        deploy = client.get(f"/api/v1/deploy?uuid={app_uuid}&force=true")
        deployment_uuid = deploy["deployments"][0]["deployment_uuid"]
        print(f"deployment_uuid={deployment_uuid}")
        last_status = ""
        for _ in range(90):
            deployment = client.get(f"/api/v1/deployments/{deployment_uuid}")
            status = str(deployment.get("status", ""))
            if status != last_status:
                print(f"status={status} updated={deployment.get('updated_at')}")
                last_status = status
            if any(done in status for done in TERMINAL_STATUSES):
                break
            time.sleep(10)

        root = requests.get(args.domain, timeout=60)
        print(f"root_status={root.status_code} content_type={root.headers.get('content-type')}")
        root.raise_for_status()
        if args.api_path:
            api_url = f"{args.domain.rstrip('/')}/{args.api_path.lstrip('/')}"
            api_response = requests.get(api_url, timeout=60)
            print(
                f"api_status={api_response.status_code} path={args.api_path} "
                f"content_type={api_response.headers.get('content-type')}"
            )
            api_response.raise_for_status()

    print(
        json.dumps(
            {
                "project_uuid": project_uuid,
                "application_uuid": app_uuid,
                "deployment_uuid": deployment_uuid,
                "domain": args.domain,
                "deployment_status": deployment.get("status") if deployment else None,
            },
            separators=(",", ":"),
        )
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
