#!/usr/bin/env bash
set -euo pipefail

skip_tests=0
skip_skill_validation=0
configuration="Release"

while [[ $# -gt 0 ]]; do
  case "$1" in
    --skip-tests)
      skip_tests=1
      shift
      ;;
    --skip-skill-validation)
      skip_skill_validation=1
      shift
      ;;
    --configuration)
      configuration="$2"
      shift 2
      ;;
    *)
      printf 'Argumento desconhecido: %s\n' "$1" >&2
      exit 2
      ;;
  esac
done

if [[ "$configuration" != "Debug" && "$configuration" != "Release" ]]; then
  printf 'Configuracao invalida: %s\n' "$configuration" >&2
  exit 2
fi

status() {
  printf '[pptx-utils] %s\n' "$1"
}

if ! command -v dotnet >/dev/null 2>&1; then
  printf 'SDK do .NET nao encontrado no PATH.\n' >&2
  exit 1
fi

script_dir="$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd)"
skill_root="$(cd -- "${script_dir}/.." && pwd)"
tool_project="${skill_root}/src/PptxOpenXmlTools/PptxOpenXmlTools.csproj"
test_project="${skill_root}/src/PptxOpenXmlTools.Tests/PptxOpenXmlTools.Tests.csproj"
package_dir="${skill_root}/bin/pptx-utils"

status "dotnet encontrado em $(command -v dotnet)"

if [[ ! -f "$tool_project" ]]; then
  printf 'Projeto CLI nao encontrado: %s\n' "$tool_project" >&2
  exit 1
fi

status 'restore: PptxOpenXmlTools'
dotnet restore "$tool_project"

status "build: PptxOpenXmlTools [$configuration]"
dotnet build "$tool_project" --configuration "$configuration" --no-restore

if [[ "$skip_tests" -eq 0 ]]; then
  if [[ ! -f "$test_project" ]]; then
    printf 'Projeto de testes nao encontrado: %s\n' "$test_project" >&2
    exit 1
  fi

  status "test: PptxOpenXmlTools.Tests [$configuration]"
  dotnet test "$test_project" --configuration "$configuration" --no-restore
else
  status 'Testes automaticos ignorados por parametro.'
fi

status "publish: PptxOpenXmlTools [$configuration] -> $package_dir"
dotnet publish "$tool_project" --configuration "$configuration" --no-restore --output "$package_dir"

executable="${package_dir}/pptx-utils"
if [[ ! -x "$executable" ]]; then
  if [[ -f "${package_dir}/pptx-utils.exe" ]]; then
    executable="${package_dir}/pptx-utils.exe"
  else
    printf 'Binario publicado nao encontrado em: %s\n' "$package_dir" >&2
    exit 1
  fi
fi

status 'smoke: pptx-utils --help'
"$executable" --help

if [[ "$skip_skill_validation" -eq 0 ]]; then
  skill_creator="${HOME}/.codex/skills/.system/skill-creator/scripts/quick_validate.py"
  if [[ -f "$skill_creator" ]]; then
    python_bin="$(command -v python3 || command -v python || true)"
    if [[ -z "$python_bin" ]]; then
      printf 'Python nao encontrado para executar quick_validate.py.\n' >&2
      exit 1
    fi

    status 'skill validation: quick_validate.py'
    "$python_bin" "$skill_creator" "$skill_root"
  else
    status 'skill-creator nao encontrado; quick_validate.py ignorado.'
  fi
else
  status 'Validacao da skill ignorada por parametro.'
fi

status "Instalacao concluida: $executable"
