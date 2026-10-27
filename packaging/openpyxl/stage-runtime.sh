#!/usr/bin/env bash
set -euo pipefail

usage() {
  echo "Usage: $0 <product-id> <isolated-runtime-root> <deb-staging-root>" >&2
  echo "Expected runtime entry: <isolated-runtime-root>/bin/python3" >&2
}

fail() {
  echo "openpyxl runtime staging: $*" >&2
  exit 1
}

[[ $# -eq 3 ]] || { usage; exit 2; }
product_id=$1
runtime_source=$(realpath -e -- "$2")
staging_root=$(realpath -e -- "$3")

[[ "$product_id" =~ ^[a-z0-9][a-z0-9-]*$ ]] || fail "invalid product id: $product_id"
[[ -d "$runtime_source" ]] || fail "runtime source is not a directory"
[[ -d "$staging_root" ]] || fail "deb staging root must already exist"
[[ -f "$runtime_source/bin/python3" && -x "$runtime_source/bin/python3" ]] || fail "runtime must contain executable bin/python3"

source_root=$(realpath -e -- "$(dirname -- "$0")/../..")
worker_source="$source_root/crates/excel-editor/python/openpyxl_worker.py"
[[ -f "$worker_source" ]] || fail "worker source is missing: $worker_source"
expected_python_version=$(tr -d '\r\n' < "$source_root/packaging/openpyxl/python-version.txt")

product_libexec="$staging_root/usr/libexec/$product_id"
[[ ! -L "$product_libexec" ]] || fail "product libexec path must not be a symlink"
install -d -- "$product_libexec"
destination="$product_libexec/openpyxl"
[[ ! -e "$destination" ]] || fail "destination already exists: $destination"
temporary=$(mktemp -d -- "$product_libexec/.openpyxl-stage.XXXXXX")
trap 'rm -rf -- "$temporary"' EXIT
install -d -- "$temporary/worker"
cp -a -- "$runtime_source" "$temporary/runtime"
install -m 0644 -- "$worker_source" "$temporary/worker/openpyxl_worker.py"

python="$temporary/runtime/bin/python3"
resolved_python=$(realpath -e -- "$python")
case "$resolved_python" in
  "$temporary/runtime"/*) ;;
  *) fail "runtime interpreter resolves outside the bundled runtime tree: $resolved_python" ;;
esac
actual_python_version=$("$python" -I -c 'import sys; print(sys.version.split()[0])')
[[ "$actual_python_version" == "$expected_python_version" ]] || fail "expected Python $expected_python_version, got $actual_python_version"

"$python" -I -c 'import openpyxl; print("openpyxl", openpyxl.__version__)' || fail "isolated runtime cannot import openpyxl"
"$python" -I "$temporary/worker/openpyxl_worker.py" --help >/dev/null || fail "worker cannot start in isolated mode"

while IFS= read -r -d '' link; do
  resolved_link=$(realpath -e -- "$link") || fail "broken runtime symlink: $link"
  case "$resolved_link" in
    "$temporary/runtime"/*) ;;
    *) fail "runtime symlink escapes the bundled tree: $link -> $resolved_link" ;;
  esac
done < <(find "$destination/runtime" -type l -print0)

if command -v ldd >/dev/null 2>&1; then
  while IFS= read -r -d '' object; do
    dependencies=$(ldd "$object" 2>&1) || fail "ldd could not inspect $object: $dependencies"
    if grep -q 'not found' <<<"$dependencies"; then
      fail "unresolved dynamic dependency in $object: $dependencies"
    fi
    while IFS= read -r dependency_line; do
      dependency_name=${dependency_line%% *}
      dependency_path=
      if [[ "$dependency_line" == *" => "* ]]; then
        dependency_path=${dependency_line#* => }
        dependency_path=${dependency_path%% *}
      elif [[ "$dependency_line" == /* ]]; then
        dependency_path=${dependency_line%% *}
        dependency_name=${dependency_path##*/}
      fi
      [[ -n "$dependency_path" ]] || continue
      resolved_dependency=$(realpath -e -- "$dependency_path") || fail "missing dependency path: $dependency_path"
      case "$resolved_dependency" in
        "$temporary/runtime"/*) ;;
        *)
          case "$dependency_name" in
            libc.so.*|libm.so.*|libpthread.so.*|libdl.so.*|librt.so.*|libutil.so.*|libresolv.so.*|ld-linux-x86-64.so.*) ;;
            *) fail "non-glibc dependency is outside the bundled runtime: $dependency_name -> $resolved_dependency" ;;
          esac
          ;;
      esac
    done <<<"$dependencies"

    if command -v objdump >/dev/null 2>&1; then
      while IFS= read -r version; do
        version=${version#GLIBC_}
        major=${version%%.*}
        minor=${version#*.}
        minor=${minor%%.*}
        if (( major > 2 || (major == 2 && minor > 28) )); then
          fail "glibc symbol $version exceeds the 2.28 target in $object"
        fi
      done < <(objdump -T "$object" 2>/dev/null | grep -oE 'GLIBC_[0-9]+\.[0-9]+' | sort -Vu)
    fi
  done < <(find "$temporary/runtime" -type f \( -name 'python3' -o -name 'python3.*' -o -name '*.so' -o -name '*.so.*' \) -print0)
fi

[[ ! -e "$destination" ]] || fail "destination appeared during staging: $destination"
mv -T -- "$temporary" "$destination"
trap - EXIT
echo "staged openpyxl worker for $product_id at $destination"
echo "runtime version: $actual_python_version"
echo "openpyxl version: $("$python" -I -c 'import openpyxl; print(openpyxl.__version__)')"
echo "Check GLIBC symbol versions and test this staged tree on glibc 2.28 and 2.31 before publishing a deb."
