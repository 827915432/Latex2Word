#!/usr/bin/env bash
# -----------------------------------------------------------------------------
# convert_with_pandoc.sh
#
# 功能概述
# --------
# 本脚本用于在“规范化工作副本”基础上，调用 Pandoc 完成 LaTeX -> Word (.docx)
# 的主体转换。
#
# 它解决的问题包括：
# 1. 检查 Pandoc / Python / reference.docx 等主转换所需依赖是否可用；
# 2. 读取 normalization-report.json，确认规范化阶段没有失败；
# 3. 自动确定主 TeX 文件、输出 docx 路径、参考样式模板路径；
# 4. 自动收集 bibliography 文件；
# 5. 自动构造 Pandoc 所需的 --resource-path；
# 6. 执行 Pandoc 主转换并保存完整日志；
# 7. 生成 JSON + Markdown 两份转换报告。
#
# 设计边界
# --------
# 本脚本只负责“主体转换”，不做以下事情：
# - 不修改原始工程；
# - 不回写规范化前的源工程；
# - 不做 docx 质量验收；
# - 不生成人工修复清单；
# - 不替代 postcheck_docx.py 的职责。
#
# 依赖
# ----
# - Bash（在 Windows 11 下可通过 Git Bash / MSYS2 / WSL 运行）
# - Pandoc
# - Python 3（优先 python3，回退 python）
#
# 典型用法
# --------
# bash scripts/convert_with_pandoc.sh --work-root /path/to/project__latex_to_word_work
#
# 或显式指定输出：
# bash scripts/convert_with_pandoc.sh \
#   --work-root /path/to/project__latex_to_word_work \
#   --output-docx /path/to/project__latex_to_word_work/output.docx
#
# 输出
# ----
# 默认在 work-root 下生成：
# - output.docx
# - pandoc-conversion.log
# - pandoc-conversion-report.json
# - pandoc-conversion-report.md
#
# 退出码
# ------
# - 0: PASS 或 PASS_WITH_WARNINGS
# - 1: FAIL
# -----------------------------------------------------------------------------

set -Eeuo pipefail

# -----------------------------------------------------------------------------
# 通用输出函数
# -----------------------------------------------------------------------------

log_info() {
  printf '[INFO] %s\n' "$*"
}

log_warn() {
  printf '[WARN] %s\n' "$*" >&2
}

log_error() {
  printf '[ERROR] %s\n' "$*" >&2
}

# -----------------------------------------------------------------------------
# 工具函数：打印用法
# -----------------------------------------------------------------------------

print_usage() {
  cat <<'EOF'
Usage:
  bash scripts/convert_with_pandoc.sh --work-root <path> [options]

Required:
  --work-root <path>              规范化工作目录。应包含 normalization-report.json。

Optional:
  --main-tex <path>               显式指定主 TeX 文件。默认从 normalization-report.json 读取。
  --normalization-json <path>     规范化报告 JSON 路径。默认 <work-root>/normalization-report.json
  --output-docx <path>            输出 docx 路径。默认 <work-root>/output.docx
  --reference-doc <path>          Word 样式模板路径。默认 <skill-root>/templates/reference.docx
  --log-file <path>               Pandoc 日志路径。默认 <work-root>/pandoc-conversion.log
  --report-json <path>            转换报告 JSON 路径。默认 <work-root>/pandoc-conversion-report.json
  --report-md <path>              转换报告 Markdown 路径。默认 <work-root>/pandoc-conversion-report.md
  -h, --help                      显示帮助

Notes:
  1. 本脚本假定上一步 normalize_tex.py 已成功执行。
  2. 若 normalization-report.json 的状态为 FAIL，本脚本会直接中止。
  3. 本脚本不会执行 docx 后检查；那是 postcheck_docx.py 的职责。
EOF
}

# -----------------------------------------------------------------------------
# 工具函数：解析脚本所在 skill 根目录
# -----------------------------------------------------------------------------

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd -P)"
SKILL_ROOT="$(cd "${SCRIPT_DIR}/.." && pwd -P)"

# -----------------------------------------------------------------------------
# 工具函数：定位 Python 命令
#
# 说明：
# - 在不同环境中，Python 命令名可能是 python3 或 python；
# - 为了避免引入 jq 依赖，本脚本用 Python 处理 JSON 与资源扫描；
# - 若 Python 不可用，则直接中止。
# -----------------------------------------------------------------------------

resolve_python_cmd() {
  if command -v python3 >/dev/null 2>&1; then
    printf 'python3'
    return 0
  fi
  if command -v python >/dev/null 2>&1; then
    printf 'python'
    return 0
  fi
  return 1
}

# -----------------------------------------------------------------------------
# 参数解析
# -----------------------------------------------------------------------------

WORK_ROOT=""
MAIN_TEX=""
NORMALIZATION_JSON=""
OUTPUT_DOCX=""
REFERENCE_DOC=""
LOG_FILE=""
REPORT_JSON=""
REPORT_MD=""

while [[ $# -gt 0 ]]; do
  case "$1" in
    --work-root)
      WORK_ROOT="${2:-}"
      shift 2
      ;;
    --main-tex)
      MAIN_TEX="${2:-}"
      shift 2
      ;;
    --normalization-json)
      NORMALIZATION_JSON="${2:-}"
      shift 2
      ;;
    --output-docx)
      OUTPUT_DOCX="${2:-}"
      shift 2
      ;;
    --reference-doc)
      REFERENCE_DOC="${2:-}"
      shift 2
      ;;
    --log-file)
      LOG_FILE="${2:-}"
      shift 2
      ;;
    --report-json)
      REPORT_JSON="${2:-}"
      shift 2
      ;;
    --report-md)
      REPORT_MD="${2:-}"
      shift 2
      ;;
    -h|--help)
      print_usage
      exit 0
      ;;
    *)
      log_error "未知参数: $1"
      print_usage
      exit 1
      ;;
  esac
done

if [[ -z "${WORK_ROOT}" ]]; then
  log_error "缺少必需参数 --work-root"
  print_usage
  exit 1
fi

# -----------------------------------------------------------------------------
# 依赖检查
# -----------------------------------------------------------------------------

if ! command -v pandoc >/dev/null 2>&1; then
  log_error "未找到 pandoc，请先安装并确保其已加入 PATH。"
  exit 1
fi

if ! PYTHON_CMD="$(resolve_python_cmd)"; then
  log_error "未找到 Python，请先安装 Python 3 并确保 python3 或 python 可用。"
  exit 1
fi

# -----------------------------------------------------------------------------
# 路径规范化与默认值填充
#
# 说明：
# - 这里不依赖 realpath，以减少跨平台差异；
# - 改用 Python 统一解析路径，避免 Windows / Bash 混合环境下路径分隔问题。
# -----------------------------------------------------------------------------

normalize_path() {
  local raw_path="$1"
  "${PYTHON_CMD}" - "$raw_path" <<'PY'
from pathlib import Path
import sys
print(str(Path(sys.argv[1]).resolve()))
PY
}

WORK_ROOT="$(normalize_path "${WORK_ROOT}")"

if [[ ! -d "${WORK_ROOT}" ]]; then
  log_error "无效的工作目录: ${WORK_ROOT}"
  exit 1
fi

if [[ -z "${NORMALIZATION_JSON}" ]]; then
  NORMALIZATION_JSON="${WORK_ROOT}/normalization-report.json"
fi
NORMALIZATION_JSON="$(normalize_path "${NORMALIZATION_JSON}")"

if [[ -z "${OUTPUT_DOCX}" ]]; then
  OUTPUT_DOCX="${WORK_ROOT}/output.docx"
fi
OUTPUT_DOCX="$(normalize_path "${OUTPUT_DOCX}")"

if [[ -z "${REFERENCE_DOC}" ]]; then
  REFERENCE_DOC="${SKILL_ROOT}/templates/reference.docx"
fi
REFERENCE_DOC="$(normalize_path "${REFERENCE_DOC}")"

if [[ -z "${LOG_FILE}" ]]; then
  LOG_FILE="${WORK_ROOT}/pandoc-conversion.log"
fi
LOG_FILE="$(normalize_path "${LOG_FILE}")"

if [[ -z "${REPORT_JSON}" ]]; then
  REPORT_JSON="${WORK_ROOT}/pandoc-conversion-report.json"
fi
REPORT_JSON="$(normalize_path "${REPORT_JSON}")"

if [[ -z "${REPORT_MD}" ]]; then
  REPORT_MD="${WORK_ROOT}/pandoc-conversion-report.md"
fi
REPORT_MD="$(normalize_path "${REPORT_MD}")"

# -----------------------------------------------------------------------------
# 规范化报告存在性检查
# -----------------------------------------------------------------------------

if [[ ! -f "${NORMALIZATION_JSON}" ]]; then
  log_error "未找到 normalization-report.json: ${NORMALIZATION_JSON}"
  exit 1
fi

if [[ ! -f "${REFERENCE_DOC}" ]]; then
  log_error "未找到 reference.docx: ${REFERENCE_DOC}"
  exit 1
fi

# -----------------------------------------------------------------------------
# 使用 Python 读取 normalization-report.json，并预生成 Pandoc 所需元数据。
#
# 输出的中间文件：
# - <work-root>/.pandoc_metadata.json
# - <work-root>/.pandoc_resource_dirs.txt
# - <work-root>/.pandoc_bibliographies.txt
#
# 这样做的好处：
# 1. Bash 不擅长解析复杂 JSON；
# 2. 不引入 jq；
# 3. 便于后续脚本或人工调试直接查看中间结果。
# -----------------------------------------------------------------------------

METADATA_JSON="${WORK_ROOT}/.pandoc_metadata.json"
RESOURCE_DIRS_TXT="${WORK_ROOT}/.pandoc_resource_dirs.txt"
BIBS_TXT="${WORK_ROOT}/.pandoc_bibliographies.txt"

"${PYTHON_CMD}" - "${WORK_ROOT}" "${NORMALIZATION_JSON}" "${MAIN_TEX}" "${METADATA_JSON}" "${RESOURCE_DIRS_TXT}" "${BIBS_TXT}" <<'PY'
import json
import re
import sys
from pathlib import Path

work_root = Path(sys.argv[1]).resolve()
normalization_json = Path(sys.argv[2]).resolve()
main_tex_cli = sys.argv[3].strip()
metadata_json = Path(sys.argv[4]).resolve()
resource_dirs_txt = Path(sys.argv[5]).resolve()
bibs_txt = Path(sys.argv[6]).resolve()

report = json.loads(normalization_json.read_text(encoding="utf-8"))

status = report.get("status")
can_continue = bool(report.get("can_continue", False))
normalized_main_tex = report.get("normalized_main_tex")
processed_tex_files = report.get("tex_files_processed") or []

if status == "FAIL" or not can_continue:
    raise SystemExit("Normalization status is FAIL or can_continue is false.")

# 确定主文件路径：
# 优先级：
# 1. 命令行 --main-tex
# 2. normalization-report.json 的 normalized_main_tex
if main_tex_cli:
    main_tex_path = Path(main_tex_cli)
    if not main_tex_path.is_absolute():
        main_tex_path = (work_root / main_tex_path).resolve()
else:
    if not normalized_main_tex:
        raise SystemExit("normalized_main_tex is missing in normalization report.")
    main_tex_path = (work_root / normalized_main_tex).resolve()

if not main_tex_path.exists():
    raise SystemExit(f"Main TeX file does not exist: {main_tex_path}")

# 处理 TeX 文件集合：
# 优先使用 normalization-report.json 中明确列出的处理文件。
tex_files = []
for rel in processed_tex_files:
    candidate = (work_root / rel).resolve()
    if candidate.exists() and candidate.suffix.lower() == ".tex":
        tex_files.append(candidate)

if not tex_files:
    tex_files = sorted(work_root.rglob("*.tex"))

# 收集 bibliography 文件与 thebibliography 使用情况。
addbib_pattern = re.compile(r"\\addbibresource(?:\[[^\]]*\])?\{([^}]+)\}")
bibliography_pattern = re.compile(r"\\bibliography\{([^}]+)\}")
thebibliography_pattern = re.compile(r"\\begin\{thebibliography\}")
cite_pattern = re.compile(
    r"\\(?:cite|citep|citet|parencite|textcite|autocite|footcite|supercite)\*?(?:\[[^\]]*\]){0,2}\{([^}]+)\}"
)

def split_csv(payload: str) -> list[str]:
    return [item.strip() for item in payload.split(",") if item.strip()]

def resolve_bib(base_dir: Path, target: str) -> Path | None:
    candidate = Path(target)
    if candidate.suffix:
        p = (base_dir / candidate).resolve()
        return p if p.exists() else None
    p = (base_dir / f"{target}.bib").resolve()
    return p if p.exists() else None

bibliographies: set[Path] = set()
thebibliography_used = False
cite_count = 0

for tex_file in tex_files:
    try:
        text = tex_file.read_text(encoding="utf-8")
    except Exception:
        try:
            text = tex_file.read_text(encoding="utf-8-sig")
        except Exception:
            try:
                text = tex_file.read_text(encoding="gbk")
            except Exception:
                continue

    if thebibliography_pattern.search(text):
        thebibliography_used = True

    cite_count += len(cite_pattern.findall(text))

    for match in addbib_pattern.finditer(text):
        target = match.group(1).strip()
        resolved = resolve_bib(tex_file.parent, target)
        if resolved is not None:
            bibliographies.add(resolved)

    for match in bibliography_pattern.finditer(text):
        targets = split_csv(match.group(1))
        for target in targets:
            resolved = resolve_bib(tex_file.parent, target)
            if resolved is not None:
                bibliographies.add(resolved)

# 构造 resource-path：
# 为提高图片资源命中率，这里将 work_root 下所有目录都加入搜索路径。
# 这是一个保守但实用的策略，能降低多文件工程相对路径差异带来的问题。
resource_dirs = sorted({p.resolve() for p in work_root.rglob("*") if p.is_dir()} | {work_root})

metadata = {
    "normalization_status": status,
    "normalization_can_continue": can_continue,
    "main_tex": str(main_tex_path),
    "tex_file_count": len(tex_files),
    "bibliography_file_count": len(bibliographies),
    "bibliographies": [str(p) for p in sorted(bibliographies)],
    "thebibliography_used": thebibliography_used,
    "cite_count": cite_count,
    "resource_dir_count": len(resource_dirs),
    "resource_dirs": [str(p) for p in resource_dirs],
}

metadata_json.write_text(json.dumps(metadata, ensure_ascii=False, indent=2), encoding="utf-8")
resource_dirs_txt.write_text("\n".join(str(p) for p in resource_dirs), encoding="utf-8")
bibs_txt.write_text("\n".join(str(p) for p in sorted(bibliographies)), encoding="utf-8")
PY

# -----------------------------------------------------------------------------
# 从中间元数据读取关键字段。
#
# 这里继续使用 Python，是为了避免 Bash 中的 JSON 字段解析脆弱化。
# -----------------------------------------------------------------------------

json_get() {
  local json_path="$1"
  local key="$2"
  "${PYTHON_CMD}" - "${json_path}" "${key}" <<'PY'
import json
import sys

path = sys.argv[1]
key = sys.argv[2]
data = json.loads(open(path, "r", encoding="utf-8").read())

value = data
for part in key.split("."):
    value = value[part]

if isinstance(value, bool):
    print("true" if value else "false")
else:
    print(value)
PY
}

MAIN_TEX_RESOLVED="$(json_get "${METADATA_JSON}" "main_tex")"
NORMALIZATION_STATUS="$(json_get "${METADATA_JSON}" "normalization_status")"
NORMALIZATION_CAN_CONTINUE="$(json_get "${METADATA_JSON}" "normalization_can_continue")"
BIB_FILE_COUNT="$(json_get "${METADATA_JSON}" "bibliography_file_count")"
THEBIBLIOGRAPHY_USED="$(json_get "${METADATA_JSON}" "thebibliography_used")"
CITE_COUNT="$(json_get "${METADATA_JSON}" "cite_count")"
RESOURCE_DIR_COUNT="$(json_get "${METADATA_JSON}" "resource_dir_count")"

if [[ "${NORMALIZATION_STATUS}" == "FAIL" || "${NORMALIZATION_CAN_CONTINUE}" != "true" ]]; then
  log_error "规范化阶段状态不可继续：status=${NORMALIZATION_STATUS}, can_continue=${NORMALIZATION_CAN_CONTINUE}"
  exit 1
fi

if [[ ! -f "${MAIN_TEX_RESOLVED}" ]]; then
  log_error "规范化后的主 TeX 文件不存在：${MAIN_TEX_RESOLVED}"
  exit 1
fi

# 若存在 cite，但既没有 bib 文件，也没有 thebibliography，则主体转换不应继续伪成功。
if [[ "${CITE_COUNT}" -gt 0 && "${BIB_FILE_COUNT}" -eq 0 && "${THEBIBLIOGRAPHY_USED}" != "true" ]]; then
  log_error "检测到文内引用，但未发现 bibliography 资源，也未发现 thebibliography 环境。"
  log_error "请先修复引用系统，再执行 Pandoc 主转换。"
  exit 1
fi

# -----------------------------------------------------------------------------
# 读取 bibliography 列表与 resource-path 列表。
# -----------------------------------------------------------------------------

mapfile -t RESOURCE_DIRS < "${RESOURCE_DIRS_TXT}"
mapfile -t BIB_FILES < "${BIBS_TXT}"

# 构造 Pandoc 的 --resource-path 字符串。
# 注意：
# - Pandoc 的 resource-path 是单个参数，其值为路径列表；
# - 在 Bash 下路径分隔符使用 ':'；
# - 对于 Git Bash / MSYS2 / WSL 下的 POSIX 风格路径，这通常是可行的。
RESOURCE_PATH=""
if [[ "${#RESOURCE_DIRS[@]}" -gt 0 ]]; then
  RESOURCE_PATH="$(IFS=:; printf '%s' "${RESOURCE_DIRS[*]}")"
else
  RESOURCE_PATH="${WORK_ROOT}"
fi

# -----------------------------------------------------------------------------
# 构造 Pandoc 命令。
#
# 参数选择说明：
# - --from=latex / --to=docx：明确指定主转换方向；
# - --standalone：生成完整文档结构；
# - --reference-doc：使用指定 Word 样式模板；
# - --resource-path：提高图片和其他资源搜索成功率；
# - --number-sections：标题编号尽量由 Pandoc/Word 结构保持；
# - --toc：生成目录基础；
# - --citeproc + --bibliography：若存在 .bib 文件则启用参考文献处理；
# - -M link-citations=true：尽量保留 citation 跳转关系。
#
# 这里不额外加入更多“花哨参数”，避免超出当前 skill 的必要设计边界。
# -----------------------------------------------------------------------------

PANDOC_CMD=(
  pandoc
  "${MAIN_TEX_RESOLVED}"
  --from=latex
  --to=docx
  --standalone
  --output="${OUTPUT_DOCX}"
  --reference-doc="${REFERENCE_DOC}"
  --resource-path="${RESOURCE_PATH}"
  --number-sections
  --toc
)

if [[ "${#BIB_FILES[@]}" -gt 0 ]]; then
  PANDOC_CMD+=(--citeproc -M "link-citations=true")
  for bib_file in "${BIB_FILES[@]}"; do
    PANDOC_CMD+=(--bibliography="${bib_file}")
  done
fi

# -----------------------------------------------------------------------------
# 将命令文本化，写入日志头部，便于调试与复现实验。
# -----------------------------------------------------------------------------

mkdir -p "$(dirname "${LOG_FILE}")"
{
  printf '=== Pandoc Conversion Log ===\n'
  printf 'Skill root: %s\n' "${SKILL_ROOT}"
  printf 'Work root: %s\n' "${WORK_ROOT}"
  printf 'Main TeX: %s\n' "${MAIN_TEX_RESOLVED}"
  printf 'Output DOCX: %s\n' "${OUTPUT_DOCX}"
  printf 'Reference DOCX: %s\n' "${REFERENCE_DOC}"
  printf 'Normalization JSON: %s\n' "${NORMALIZATION_JSON}"
  printf 'Bibliography file count: %s\n' "${BIB_FILE_COUNT}"
  printf 'Resource dir count: %s\n' "${RESOURCE_DIR_COUNT}"
  printf '\nCommand:\n'
  printf '  '
  printf '%q ' "${PANDOC_CMD[@]}"
  printf '\n\n=== Pandoc stdout/stderr ===\n'
} > "${LOG_FILE}"

# -----------------------------------------------------------------------------
# 执行 Pandoc 主转换。
#
# 说明：
# - 使用 pipefail + tee，同时把 stdout/stderr 输出到屏幕与日志；
# - 保留 Pandoc 的真实退出码；
# - 不在这里做“转换是否合格”的最终判断，那是 report + postcheck 的职责；
# - 这里只判断主转换层面成功或失败。
# -----------------------------------------------------------------------------

log_info "开始执行 Pandoc 主转换..."
set +e
(
  cd "${WORK_ROOT}" || exit 1
  "${PANDOC_CMD[@]}"
) 2>&1 | tee -a "${LOG_FILE}"
PANDOC_EXIT_CODE="${PIPESTATUS[0]}"
set -e

# -----------------------------------------------------------------------------
# 根据 Pandoc 退出码和输出文件存在性，形成主体转换结论。
# -----------------------------------------------------------------------------

CONVERSION_STATUS="PASS"
CAN_CONTINUE="true"
WARNING_MESSAGES=()
FAILURE_REASON=""

if [[ "${PANDOC_EXIT_CODE}" -ne 0 ]]; then
  CONVERSION_STATUS="FAIL"
  CAN_CONTINUE="false"
  FAILURE_REASON="pandoc_exit_nonzero"
fi

if [[ ! -f "${OUTPUT_DOCX}" ]]; then
  CONVERSION_STATUS="FAIL"
  CAN_CONTINUE="false"
  FAILURE_REASON="output_docx_missing"
fi

if [[ -f "${OUTPUT_DOCX}" && ! -s "${OUTPUT_DOCX}" ]]; then
  CONVERSION_STATUS="FAIL"
  CAN_CONTINUE="false"
  FAILURE_REASON="output_docx_empty"
fi

# 若 Pandoc 转换成功，但未启用 bibliography 处理且工程中存在 cite，则作为告警，而非静默忽略。
if [[ "${CONVERSION_STATUS}" != "FAIL" && "${CITE_COUNT}" -gt 0 && "${BIB_FILE_COUNT}" -eq 0 && "${THEBIBLIOGRAPHY_USED}" == "true" ]]; then
  CONVERSION_STATUS="PASS_WITH_WARNINGS"
  WARNING_MESSAGES+=("检测到 thebibliography 环境；Pandoc 未启用 .bib 驱动的 citeproc，参考文献与链接效果需在后检查阶段重点核对。")
fi

# 若没有 bibliography 文件，但也没有 cite，则不是错误。
if [[ "${CONVERSION_STATUS}" != "FAIL" && "${BIB_FILE_COUNT}" -eq 0 && "${CITE_COUNT}" -eq 0 ]]; then
  WARNING_MESSAGES+=("未检测到 .bib bibliography 文件；若文档确实不使用引用系统，这不是问题。")
  if [[ "${#WARNING_MESSAGES[@]}" -gt 0 && "${CONVERSION_STATUS}" == "PASS" ]]; then
    CONVERSION_STATUS="PASS_WITH_WARNINGS"
  fi
fi

# 若资源目录特别多，提示用户后续若出现性能问题应检查工程结构。
if [[ "${RESOURCE_DIR_COUNT}" -gt 200 ]]; then
  WARNING_MESSAGES+=("resource-path 包含的目录较多；若后续性能较差，可考虑收敛资源目录结构。")
  if [[ "${CONVERSION_STATUS}" == "PASS" ]]; then
    CONVERSION_STATUS="PASS_WITH_WARNINGS"
  fi
fi

# -----------------------------------------------------------------------------
# 生成 JSON 与 Markdown 报告。
#
# 说明：
# - Bash 不适合优雅生成复杂 JSON/Markdown；
# - 因此继续使用 Python 把结果结构化输出；
# - 这一步只记录“Pandoc 主转换”事实，不代替 postcheck。
# -----------------------------------------------------------------------------

# 将 warning 列表编码为 JSON 数组文本，避免 shell 转义问题。
WARNINGS_JSON="$("${PYTHON_CMD}" - <<'PY' "${WARNING_MESSAGES[@]}"
import json
import sys
print(json.dumps(sys.argv[1:], ensure_ascii=False))
PY
)"

"${PYTHON_CMD}" - \
  "${REPORT_JSON}" \
  "${REPORT_MD}" \
  "${WORK_ROOT}" \
  "${MAIN_TEX_RESOLVED}" \
  "${OUTPUT_DOCX}" \
  "${REFERENCE_DOC}" \
  "${LOG_FILE}" \
  "${NORMALIZATION_JSON}" \
  "${METADATA_JSON}" \
  "${CONVERSION_STATUS}" \
  "${CAN_CONTINUE}" \
  "${PANDOC_EXIT_CODE}" \
  "${FAILURE_REASON}" \
  "${WARNINGS_JSON}" <<'PY'
import json
import sys
from pathlib import Path

report_json = Path(sys.argv[1]).resolve()
report_md = Path(sys.argv[2]).resolve()
work_root = Path(sys.argv[3]).resolve()
main_tex = Path(sys.argv[4]).resolve()
output_docx = Path(sys.argv[5]).resolve()
reference_doc = Path(sys.argv[6]).resolve()
log_file = Path(sys.argv[7]).resolve()
normalization_json = Path(sys.argv[8]).resolve()
metadata_json = Path(sys.argv[9]).resolve()
status = sys.argv[10]
can_continue = sys.argv[11].lower() == "true"
pandoc_exit_code = int(sys.argv[12])
failure_reason = sys.argv[13]
warnings = json.loads(sys.argv[14])

metadata = json.loads(metadata_json.read_text(encoding="utf-8"))
normalization = json.loads(normalization_json.read_text(encoding="utf-8"))

report = {
    "status": status,
    "can_continue": can_continue,
    "work_root": str(work_root),
    "main_tex": str(main_tex),
    "output_docx": str(output_docx),
    "reference_doc": str(reference_doc),
    "normalization_json": str(normalization_json),
    "pandoc_log": str(log_file),
    "pandoc_exit_code": pandoc_exit_code,
    "failure_reason": failure_reason if failure_reason else None,
    "metrics": {
        "bibliography_file_count": metadata.get("bibliography_file_count", 0),
        "resource_dir_count": metadata.get("resource_dir_count", 0),
        "tex_file_count": metadata.get("tex_file_count", 0),
        "cite_count": metadata.get("cite_count", 0),
    },
    "inventory": {
        "bibliographies": metadata.get("bibliographies", []),
        "resource_dirs": metadata.get("resource_dirs", []),
        "thebibliography_used": metadata.get("thebibliography_used", False),
    },
    "summary": {
        "normalization_status": normalization.get("status"),
        "normalization_can_continue": normalization.get("can_continue"),
        "pandoc_exit_code": pandoc_exit_code,
        "output_exists": output_docx.exists(),
        "output_size_bytes": output_docx.stat().st_size if output_docx.exists() else 0,
    },
    "warnings": warnings,
    "recommendations": [],
}

if status == "FAIL":
    report["recommendations"].append("先查看 pandoc-conversion.log 中的错误信息，再修复阻塞问题后重试。")
else:
    report["recommendations"].append("可进入 postcheck_docx.py 阶段，对生成的 docx 做结构级后检查。")

if metadata.get("bibliography_file_count", 0) > 0:
    report["recommendations"].append("参考文献已作为 Pandoc 主转换输入；仍应在后检查阶段核对引用跳转与样式。")

if warnings:
    report["recommendations"].append("存在转换告警；请在后检查和人工修复阶段重点核对这些对象。")

report_json.parent.mkdir(parents=True, exist_ok=True)
report_json.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")

lines = []
lines.append("# Pandoc Conversion Report")
lines.append("")
lines.append(f"- Status: **{report['status']}**")
lines.append(f"- Can continue: **{report['can_continue']}**")
lines.append(f"- Work root: `{report['work_root']}`")
lines.append(f"- Main TeX: `{report['main_tex']}`")
lines.append(f"- Output DOCX: `{report['output_docx']}`")
lines.append(f"- Reference DOCX: `{report['reference_doc']}`")
lines.append(f"- Pandoc log: `{report['pandoc_log']}`")
lines.append("")

lines.append("## Summary")
lines.append("")
for key, value in report["summary"].items():
    lines.append(f"- {key}: **{value}**")
lines.append("")

lines.append("## Metrics")
lines.append("")
for key, value in report["metrics"].items():
    lines.append(f"- {key}: **{value}**")
lines.append("")

lines.append("## Warnings")
lines.append("")
if report["warnings"]:
    for item in report["warnings"]:
        lines.append(f"- {item}")
else:
    lines.append("- (none)")
lines.append("")

lines.append("## Recommendations")
lines.append("")
if report["recommendations"]:
    for item in report["recommendations"]:
        lines.append(f"- {item}")
else:
    lines.append("- (none)")
lines.append("")

report_md.parent.mkdir(parents=True, exist_ok=True)
report_md.write_text("\n".join(lines), encoding="utf-8")
PY

# -----------------------------------------------------------------------------
# 控制台摘要输出
# -----------------------------------------------------------------------------

log_info "Pandoc 主转换结束。"
log_info "Status: ${CONVERSION_STATUS}"
log_info "Main TeX: ${MAIN_TEX_RESOLVED}"
log_info "Output DOCX: ${OUTPUT_DOCX}"
log_info "Log file: ${LOG_FILE}"
log_info "JSON report: ${REPORT_JSON}"
log_info "Markdown report: ${REPORT_MD}"

if [[ "${CONVERSION_STATUS}" == "FAIL" ]]; then
  log_error "Pandoc 主转换失败。请先检查日志：${LOG_FILE}"
  exit 1
fi

exit 0