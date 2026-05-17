---
name: coolify-deploy
description: Automate and validate Coolify deployments through the Coolify API using Python requests for GitHub-backed Docker Compose apps. Use when the user asks to publish, deploy, create a Coolify project/environment/application, configure runtime env vars, trigger a deployment, inspect deployment status/logs, or validate a public domain after deploy, especially with deploy_coolify.md, docker-compose.yml, Dockerfile, .env.example, and COOLIFY_API_KEY.
---

# Coolify Deploy

## Workflow

1. Read the repo deploy notes first, usually `deploy_coolify.md`, then inspect `docker-compose.yml`, Dockerfile, entrypoint, `.env.example`, and `git status`.
2. Verify locally before touching Coolify: `docker compose --env-file .env.example config`, tests/builds relevant to the stack, and a container smoke test when feasible.
3. Commit and push deploy changes before asking Coolify to build. Coolify deploys from the remote branch, not local uncommitted files.
4. When the target project, environment, domain, application, and auto-deploy branch already exist and are healthy, do not trigger a manual redeploy for every code change. Push the target branch and monitor the deployment created by the GitHub App/webhook instead. Trigger deploy manually only for initial provisioning, changed Coolify configuration/env vars, disabled or failed auto-deploy, or an explicit user request.
5. Authenticate with `COOLIFY_API_KEY` from the environment. Do not print token values or generated secrets.
6. Use the Coolify API idempotently: find or create project, environment, application, env vars, then deploy by application UUID and poll deployment status.
7. Validate the public domain by direct HTTP/HTTPS requests to `/` and expected API routes. Report exact status codes and Coolify UUIDs.

## Defaults

- Default base URL: `https://coolify.drg.ink`, unless `COOLIFY_URL` is set or the user provides another URL.
- Prefer branch `master` only when it is current and pushed; otherwise use the branch the user explicitly requests.
- For Docker Compose apps, pass `docker_compose_location` with a leading slash, for example `/docker-compose.yml`.
- Use `build_pack=dockercompose`, `ports_exposes=80`, `is_auto_deploy_enabled=true`, and `is_force_https_enabled=true` unless the repo says otherwise.
- If Coolify creates an automatic `production` environment but the requested environment is `prod`, create/use `prod`; remove the empty automatic environment only when it has no resources and the runbook requires `prod`.

## Safety

- Never print `COOLIFY_API_KEY`, generated `SECRET_KEY`, database passwords, manual webhook secrets, sentinel tokens, or full application JSON that may contain secrets.
- When listing resources, emit filtered summaries: names, UUIDs, status, branch, repository, and domain only.
- Generate strong runtime secrets when placeholders such as `change_me` are present.
- Do not store Coolify API tokens in repo files. Use environment variables or one-off shell scope.
- If using web for Coolify API docs, restrict to official Coolify documentation unless the user asks otherwise.

## Python Script

Use `scripts/deploy_coolify.py` for the repeatable path. It uses `requests` and:

- reads `COOLIFY_API_KEY`;
- creates or reuses project/environment/app;
- updates env vars in bulk;
- triggers a force deploy;
- polls deployment status;
- validates root and API URLs.

Example:

```powershell
$env:COOLIFY_API_KEY = "<token>"
python C:\Users\imale\.codex\skills\coolify-deploy\scripts\deploy_coolify.py `
  --project-name ordinais `
  --environment-name prod `
  --repository im-alexandre/ordinais `
  --branch master `
  --domain https://electremor.drg.ink `
  --service-name electre_mor `
  --server-uuid uswc0ggc0wwo0g8wc0oscwwo `
  --github-app-uuid c8k8s8wcw8kksg4ws8gc8w0s `
  --compose-location /docker-compose.yml `
  --api-path /api/v1/ `
  --django-default-envs
```

Add custom envs with repeated `--env KEY=VALUE`. Use `--skip-deploy` only for provisioning without publishing.

If the script fails because a Coolify endpoint changed, inspect official API docs and retry with the smallest correction.
