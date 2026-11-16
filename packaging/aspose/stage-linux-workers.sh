#!/usr/bin/env bash
# Stage published, self-contained linux-x64 workers. Never installs host software.
set -euo pipefail
fail() { echo "Aspose Linux staging: $*" >&2; exit 1; }
is_elf() { LC_ALL=C head -c 4 "$1" | LC_ALL=C grep -q $'^\177ELF$'; }
[[ $# == 2 ]] || fail "usage: $0 <published-workers-root> <application-root>"
for command in realpath find xargs objdump head grep sed sort tail install cp chmod sha256sum; do
    command -v "$command" >/dev/null || fail "required staging tool is missing: $command"
done
source_root=$(realpath -e -- "$1")
app_root=$(realpath -e -- "$2")
[[ -d "$source_root" && -d "$app_root" ]] || fail "directories required"
[[ ! -e "$app_root/workers" ]] || fail "worker destination already exists"

for entry in aspose:AsposeWorker aspose-words:AsposeWordsWorker; do
    directory=${entry%:*}; name=${entry#*:}; source="$source_root/$directory"
    [[ -d "$source" && ! -L "$source" ]] || fail "missing directory: $directory"
    [[ -f "$source/$name" && -f "$source/$name.dll" && -f "$source/$name.runtimeconfig.json" &&
       -f "$source/libcoreclr.so" && -f "$source/libhostfxr.so" && -f "$source/libSkiaSharp.so" ]] ||
        fail "incomplete self-contained worker: $directory"
    [[ -z "$(find "$source" -type l -print -quit)" ]] || fail "symlinks are not accepted"
    [[ -z "$(find "$source" -type f \( -iname '*.exe' -o -iname '*.pdb' \) -print -quit)" ]] ||
        fail "Windows executables/debug symbols in Linux worker"
    is_elf "$source/$name" || fail "worker apphost is not ELF"
    while IFS= read -r -d '' file; do
        is_elf "$file" || continue
        objdump -f "$file" | grep -q 'architecture: i386:x86-64' || fail "not x86_64: $file"
        maximum=$(objdump -T "$file" | { grep -o 'GLIBC_[0-9.]*' || true; } |
            sed 's/GLIBC_//' | sort -V | tail -1)
        [[ "$maximum" =~ ^[0-9]+\.[0-9]+(\.[0-9]+)?$ ]] || fail "unrecognized GLIBC requirements: $file"
        comparable="$maximum"
        [[ "$comparable" =~ ^[0-9]+\.[0-9]+$ ]] && comparable="$comparable.0"
        [[ "$(printf '%s\n' "$comparable" 2.28.0 | sort -V | tail -1)" == 2.28.0 ]] ||
            fail "GLIBC $maximum exceeds 2.28: $file"
    done < <(find "$source" -type f -print0)
done

install -d -m 0755 "$app_root/workers"
for entry in aspose:AsposeWorker aspose-words:AsposeWordsWorker; do
    directory=${entry%:*}; name=${entry#*:}
    cp -a -- "$source_root/$directory" "$app_root/workers/$directory"
    # NTFS publish directories can appear as 0777 in WSL. Never ship that mode.
    find "$app_root/workers/$directory" -type d -exec chmod 0755 {} +
    find "$app_root/workers/$directory" -type f -exec chmod 0644 {} +
    chmod 0755 "$app_root/workers/$directory/$name"
done
(cd "$app_root/workers"; find aspose aspose-words -type f -print0 | sort -z | xargs -0 sha256sum > MANIFEST.sha256)
echo "Aspose Cells/Words linux-x64 self-contained workers staged; evaluation output retained"
