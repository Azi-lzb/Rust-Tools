# Isolated openpyxl runtime staging

`stage-runtime.sh` prepares the openpyxl worker subtree inside an existing deb staging root. It does not build a Python runtime, install packages, create a deb, or modify product route capabilities.

The supplied relocatable runtime must already contain `bin/python3`, openpyxl, and et-xmlfile. The intended pinned versions are Python 3.8.10 (`python-version.txt`), openpyxl 3.1.5, and et-xmlfile 2.0.0 (`requirements.txt`). The staging script rejects a different Python version, then copies it to:

```text
<deb-root>/usr/libexec/<product>/openpyxl/runtime/
<deb-root>/usr/libexec/<product>/openpyxl/worker/openpyxl_worker.py
```

This matches `OpenpyxlWorkerConfig::for_product_binary` for `/usr/bin/<product>`. The runtime must resolve its interpreter and internal symlinks within the staged runtime tree. Staging checks `python -I`, imports openpyxl, starts the worker help path, and rejects unresolved dynamic dependencies found by `ldd`.

Example:

```sh
bash packaging/openpyxl/stage-runtime.sh pytools /opt/build/python-runtime deb-staging/pytools
```

No Python 3.8.10 product runtime artifact is committed by this workspace. The worker has only been tested in this development session with Python 3.12.3/openpyxl 3.1.5; that is not a production runtime assertion. Before release, source an offline relocatable Python 3.8.10 runtime containing the pinned wheels, inspect GLIBC symbols for the interpreter and extension modules, and run the staged worker on both glibc 2.28 and glibc 2.31 systems. The current Windows/WSL synthetic tests do not prove relocatability or target Linux compatibility.
