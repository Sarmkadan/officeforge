#!/usr/bin/env python3
"""
Simple build helper used by the task‑factory.

Running this script will invoke `dotnet build` for the solution in the current
working directory and forward the output to the console.

If the .NET SDK is not installed, a clear error message is printed and the
process exits with code 1.
"""

import subprocess
import sys
from pathlib import Path

def main() -> None:
    # Ensure we are in the repository root (where a .sln or .csproj is expected)
    repo_root = Path(__file__).resolve().parent
    # Prefer a solution file if present; otherwise let `dotnet build` discover a project.
    solution_files = list(repo_root.glob("*.sln"))
    build_cmd = ["dotnet", "build"]
    if solution_files:
        # Use the first solution file found.
        build_cmd.append(str(solution_files[0]))

    try:
        result = subprocess.run(
            build_cmd,
            cwd=repo_root,
            capture_output=True,
            text=True,
        )
    except FileNotFoundError:
        sys.stderr.write("Error: The .NET SDK (dotnet) is not installed or not on PATH.\n")
        sys.exit(1)

    # Forward stdout and stderr exactly as produced by the dotnet CLI.
    sys.stdout.write(result.stdout)
    sys.stderr.write(result.stderr)
    sys.exit(result.returncode)


if __name__ == "__main__":
    main()
