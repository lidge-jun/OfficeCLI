using System.Diagnostics;
using OfficeCli.Tests.Hwpx;

namespace OfficeCli.Tests.Hwp;

public partial class HwpBridgeSidecarTests
{
    private static string LocateBridgeDll()
    {
        var dir = new DirectoryInfo(AppContext.BaseDirectory);
        while (dir != null)
        {
            var candidate = Path.Combine(
                dir.FullName,
                "src/rhwp-officecli-bridge/bin/Debug/net10.0/rhwp-officecli-bridge.dll");
            if (File.Exists(candidate)) return candidate;
            dir = dir.Parent;
        }
        throw new FileNotFoundException("rhwp-officecli-bridge.dll was not built.");
    }

    private ProcessResult RunBridge(string bridgeDll, string fakeRhwp, string[] args, string? fakeRhwpApi = null)
    {
        var psi = new ProcessStartInfo
        {
            FileName = "dotnet",
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true
        };
        psi.ArgumentList.Add(bridgeDll);
        foreach (var arg in args) psi.ArgumentList.Add(arg);
        psi.Environment["OFFICECLI_RHWP_BIN"] = fakeRhwp;
        psi.Environment["OFFICECLI_RHWP_API_BIN"] = fakeRhwpApi
            ?? Path.Combine(Path.GetTempPath(), "officecli-test-no-rhwp-api");
        using var process = Process.Start(psi)!;
        var stdout = process.StandardOutput.ReadToEnd();
        var stderr = process.StandardError.ReadToEnd();
        process.WaitForExit();
        return new ProcessResult(process.ExitCode, stdout, stderr);
    }

    private static (int ExitCode, string Stdout) InvokeOfficeCli(string[] args)
    {
        // Run the real CLI out-of-process. Console.Out/Error and environment
        // variables are process-global; invoking the root in-process races the
        // test runner's own Console capture and produced empty stdout / false
        // bridge-missing results even with xUnit parallelism disabled.
        var psi = new ProcessStartInfo
        {
            FileName = "dotnet",
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
        };
        psi.ArgumentList.Add(LocateOfficeCliDll());
        foreach (var arg in args) psi.ArgumentList.Add(arg);
        psi.Environment["OFFICECLI_RHWP_TIMEOUT_SCALE"] = "5";
        using var process = Process.Start(psi)!;
        var stdout = process.StandardOutput.ReadToEnd();
        var stderr = process.StandardError.ReadToEnd();
        process.WaitForExit();

        if (process.ExitCode != 0)
        {
            if (!string.IsNullOrWhiteSpace(stderr))
                Console.Error.WriteLine($"[officecli stderr] {stderr.Trim()}");
            if (!string.IsNullOrWhiteSpace(stdout))
                Console.Error.WriteLine($"[officecli stdout] {stdout.Trim()}");
        }

        return (process.ExitCode, stdout);
    }

    private static string LocateOfficeCliDll()
    {
        var dir = new DirectoryInfo(AppContext.BaseDirectory);
        while (dir != null)
        {
            var bin = Path.Combine(dir.FullName, "src/officecli/bin/Debug/net10.0");
            if (Directory.Exists(bin))
            {
                var candidate = Directory.GetFiles(bin, "officecli.dll", SearchOption.AllDirectories)
                    .OrderByDescending(File.GetLastWriteTimeUtc)
                    .FirstOrDefault();
                if (candidate != null) return candidate;
            }
            dir = dir.Parent;
        }
        throw new FileNotFoundException("officecli.dll was not built.");
    }

    private string CreateFakeRhwp()
    {
        var path = Path.Combine(Path.GetTempPath(), $"fake-rhwp-{Guid.NewGuid():N}");
        File.WriteAllText(path, """
#!/bin/sh
if [ "$1" = "--version" ]; then
  echo "rhwp v0.test"
  exit 0
fi
cmd="$1"
input=""
if [ "$#" -gt 1 ] && [ "${2#--}" = "$2" ]; then
  input="$2"
fi
out="output"
while [ "$#" -gt 0 ]; do
  if [ "$1" = "--input" ]; then
    shift
    input="$1"
  fi
  if [ "$1" = "--output" ] || [ "$1" = "-o" ]; then
    shift
    out="$1"
  fi
  shift
done
mkdir -p "$out"
if [ "$cmd" = "export-text" ]; then
  if [ "${input##*.}" = "hwpx" ]; then
    python3 - "$input" "$out/page_001.txt" <<'PY'
import html
import re
import sys
import zipfile

input_path, output_path = sys.argv[1], sys.argv[2]
texts = []
with zipfile.ZipFile(input_path) as archive:
    for name in sorted(archive.namelist()):
        if name.startswith("Contents/section") and name.endswith(".xml"):
            xml = archive.read(name).decode("utf-8", errors="ignore")
            texts.extend(html.unescape(value) for value in re.findall(r"<(?:\w+:)?t[^>]*>(.*?)</(?:\w+:)?t>", xml, re.S))
with open(output_path, "w", encoding="utf-8") as output:
    output.write(" ".join(texts))
PY
  else
    cat "$input" > "$out/page_001.txt"
  fi
  exit 0
fi
if [ "$cmd" = "export-svg" ]; then
  printf '<svg><text>page</text></svg>' > "$out/page.svg"
  exit 0
fi
echo "unexpected command: $cmd" >&2
exit 2
""");
        if (!OperatingSystem.IsWindows())
            File.SetUnixFileMode(path,
                UnixFileMode.UserRead | UnixFileMode.UserWrite | UnixFileMode.UserExecute);
        _tempPaths.Add(path);
        return path;
    }

    private string CreateFakeRhwpApi()
    {
        var path = Path.Combine(Path.GetTempPath(), $"fake-rhwp-api-{Guid.NewGuid():N}");
        File.WriteAllText(path, """
#!/bin/sh
cmd="$1"
if [ "$cmd" = "--help" ] || [ "$cmd" = "-h" ]; then
  echo "rhwp-field-bridge create-blank|read-text|render-svg|render-png|export-pdf|export-markdown|document-info|diagnostics|dump-controls|dump-pages|thumbnail|list-fields|get-field|set-field|fill-fields|replace-text|insert-text|get-cell-text|scan-cells|set-cell-text|convert-to-editable|native-op|save-as-hwp --format hwp|hwpx [--input <path>] [--op <native-op>] [--output <path>] [--out-dir <dir>] --json"
  exit 0
fi
if [ "$cmd" = "create-blank" ]; then
  output=""
  while [ "$#" -gt 0 ]; do
    if [ "$1" = "--output" ]; then
      shift
      output="$1"
    fi
    shift
  done
  printf 'fake blank hwp' > "$output"
  printf '{"created":true,"operation":"create-blank","output":"%s","engineVersion":"rhwp-api v0.test","format":"hwp","warnings":["experimental create-blank"]}\n' "$output"
  exit 0
fi
if [ "$cmd" = "read-text" ]; then
  input=""
  while [ "$#" -gt 0 ]; do
    if [ "$1" = "--input" ]; then
      shift
      input="$1"
    fi
    shift
  done
  text="$(cat "$input" 2>/dev/null || printf '')"
  if [ "${input##*.}" = "hwpx" ]; then
    text="$(python3 - "$input" <<'PY'
import html
import re
import sys
import zipfile

texts = []
with zipfile.ZipFile(sys.argv[1]) as archive:
    for name in sorted(archive.namelist()):
        if name.startswith("Contents/section") and name.endswith(".xml"):
            xml = archive.read(name).decode("utf-8", errors="ignore")
            texts.extend(html.unescape(value) for value in re.findall(r"<(?:\w+:)?t[^>]*>(.*?)</(?:\w+:)?t>", xml, re.S))
print(" ".join(texts), end="")
PY
)"
  fi
  python3 - "$text" <<'PY'
import json
import sys

text = sys.argv[1]
print(json.dumps({
    "text": text,
    "pages": [{"page": 1, "text": text}],
    "engineVersion": "rhwp-api v0.test",
    "format": "hwp",
    "warnings": []
}, ensure_ascii=False))
PY
  exit 0
fi
if [ "$cmd" = "render-svg" ]; then
  out=""
  while [ "$#" -gt 0 ]; do
    if [ "$1" = "--out-dir" ]; then
      shift
      out="$1"
    fi
    shift
  done
  mkdir -p "$out"
  printf '<svg><text>api page</text></svg>' > "$out/page_001.svg"
  printf '{"engineVersion":"rhwp-api v0.test","format":"hwp","manifest":"%s/manifest.json","pages":[{"page":1,"path":"%s/page_001.svg","sha256":"abc123"}],"warnings":[]}\n' "$out" "$out"
  exit 0
fi
if [ "$cmd" = "render-png" ]; then
  out=""
  while [ "$#" -gt 0 ]; do
    if [ "$1" = "--out-dir" ]; then
      shift
      out="$1"
    fi
    shift
  done
  mkdir -p "$out"
  printf 'PNG' > "$out/page_001.png"
  printf '{"engineVersion":"rhwp-api v0.test","format":"hwp","manifest":"%s/manifest.json","pages":[{"page":1,"path":"%s/page_001.png","sha256":"png123"}],"warnings":[]}\n' "$out" "$out"
  exit 0
fi
if [ "$cmd" = "export-pdf" ]; then
  output=""
  while [ "$#" -gt 0 ]; do
    if [ "$1" = "--output" ]; then
      shift
      output="$1"
    fi
    shift
  done
  printf 'PDF' > "$output"
  printf '{"pdf":{"path":"%s","bytes":3,"sha256":"pdf123"},"pages":[{"page":1}],"engineVersion":"rhwp-api v0.test","format":"hwp","warnings":[]}\n' "$output"
  exit 0
fi
if [ "$cmd" = "export-markdown" ]; then
  printf '%s\n' '{"markdown":"# Fake HWP\n\n셀값","pages":[{"page":1,"markdown":"# Fake HWP"}],"engineVersion":"rhwp-api v0.test","format":"hwp","warnings":[]}'
  exit 0
fi
if [ "$cmd" = "document-info" ]; then
  printf '%s\n' '{"info":{"pages":1,"sections":1},"engineVersion":"rhwp-api v0.test","format":"hwp","warnings":[]}'
  exit 0
fi
if [ "$cmd" = "diagnostics" ]; then
  printf '%s\n' '{"diagnostics":[],"engineVersion":"rhwp-api v0.test","format":"hwp","warnings":[]}'
  exit 0
fi
if [ "$cmd" = "dump-controls" ]; then
  printf '%s\n' '{"dump":"full control dump","engineVersion":"rhwp-api v0.test","format":"hwp","warnings":[]}'
  exit 0
fi
if [ "$cmd" = "dump-pages" ]; then
  printf '%s\n' '{"dump":"page 1 dump","pages":[{"page":1,"dump":"page 1 dump"}],"engineVersion":"rhwp-api v0.test","format":"hwp","warnings":[]}'
  exit 0
fi
if [ "$cmd" = "thumbnail" ]; then
  output=""
  while [ "$#" -gt 0 ]; do
    if [ "$1" = "--output" ]; then
      shift
      output="$1"
    fi
    shift
  done
  printf 'PNG' > "$output"
  printf '{"thumbnail":{"path":"%s","format":"png","width":120,"height":90},"engineVersion":"rhwp-api v0.test","format":"hwp","warnings":[]}\n' "$output"
  exit 0
fi
if [ "$cmd" = "list-fields" ]; then
  printf '%s\n' '{"fields":[{"fieldId":7,"fieldType":"clickhere","name":"applicant","value":"홍길동"}],"engineVersion":"rhwp-api v0.test","format":"hwp","warnings":[]}'
  exit 0
fi
if [ "$cmd" = "get-field" ]; then
  printf '%s\n' '{"field":{"ok":true,"fieldId":7,"value":"홍길동"},"engineVersion":"rhwp-api v0.test","format":"hwp","warnings":[]}'
  exit 0
fi
if [ "$cmd" = "set-field" ]; then
  output=""
  value=""
  while [ "$#" -gt 0 ]; do
    if [ "$1" = "--output" ]; then
      shift
      output="$1"
    elif [ "$1" = "--value" ]; then
      shift
      value="$1"
    fi
    shift
  done
  printf 'fake hwp output' > "$output"
  printf '{"field":{"ok":true,"fieldId":7,"oldValue":"","newValue":"%s"},"output":"%s","engineVersion":"rhwp-api v0.test","format":"hwp","warnings":["experimental set-field"]}\n' "$value" "$output"
  exit 0
fi
if [ "$cmd" = "replace-text" ]; then
  format="hwp"
  input=""
  output=""
  query=""
  value=""
  while [ "$#" -gt 0 ]; do
    if [ "$1" = "--format" ]; then
      shift
      format="$1"
    elif [ "$1" = "--input" ]; then
      shift
      input="$1"
    elif [ "$1" = "--output" ]; then
      shift
      output="$1"
    elif [ "$1" = "--query" ]; then
      shift
      query="$1"
    elif [ "$1" = "--value" ]; then
      shift
      value="$1"
    fi
    shift
  done
  if [ "$format" = "hwpx" ]; then
    python3 - "$input" "$output" "$query" "$value" <<'PY'
import sys
import zipfile

input_path, output_path, query, value = sys.argv[1], sys.argv[2], sys.argv[3], sys.argv[4]
with zipfile.ZipFile(input_path) as source, zipfile.ZipFile(output_path, "w") as target:
    for info in source.infolist():
        data = source.read(info.filename)
        if info.filename.endswith((".xml", ".hpf", ".rdf", ".txt")):
            text = data.decode("utf-8", errors="ignore")
            data = text.replace(query, value).encode("utf-8")
        target.writestr(info, data)
PY
  else
    printf '%s %s' "$value" "$value" > "$output"
  fi
  printf '{"replacement":{"ok":true,"count":2},"output":"%s","engineVersion":"rhwp-api v0.test","format":"%s","warnings":["experimental replace-text"]}\n' "$output" "$format"
  exit 0
fi
if [ "$cmd" = "insert-text" ]; then
  format="hwp"
  output=""
  value=""
  section="0"
  paragraph="0"
  offset="0"
  while [ "$#" -gt 0 ]; do
    if [ "$1" = "--format" ]; then
      shift
      format="$1"
    elif [ "$1" = "--output" ]; then
      shift
      output="$1"
    elif [ "$1" = "--value" ]; then
      shift
      value="$1"
    elif [ "$1" = "--section" ]; then
      shift
      section="$1"
    elif [ "$1" = "--paragraph" ] || [ "$1" = "--para" ]; then
      shift
      paragraph="$1"
    elif [ "$1" = "--offset" ]; then
      shift
      offset="$1"
    fi
    shift
  done
  printf '%s' "$value" > "$output"
  printf '{"inserted":true,"operation":"insert-text","output":"%s","engineVersion":"rhwp-api v0.test","format":"%s","section":%s,"paragraph":%s,"offset":%s,"warnings":["experimental insert-text"]}\n' "$output" "$format" "$section" "$paragraph" "$offset"
  exit 0
fi
if [ "$cmd" = "get-cell-text" ]; then
  printf '%s\n' '{"cell":{"section":0,"parentPara":2,"control":0,"cell":1,"cellPara":0,"offset":0,"count":1000,"text":"셀값"},"engineVersion":"rhwp-api v0.test","format":"hwp","warnings":[]}'
  exit 0
fi
if [ "$cmd" = "scan-cells" ]; then
  printf '%s\n' '{"cells":[{"section":0,"parentPara":2,"control":0,"cell":1,"cellPara":0,"text":"셀값"}],"engineVersion":"rhwp-api v0.test","format":"hwp","warnings":[]}'
  exit 0
fi
if [ "$cmd" = "convert-to-editable" ]; then
  output=""
  while [ "$#" -gt 0 ]; do
    if [ "$1" = "--output" ]; then
      shift
      output="$1"
    fi
    shift
  done
  printf 'fake editable hwp' > "$output"
  printf '{"converted":{"ok":true,"converted":true},"output":"%s","outputFormat":"hwp","engineVersion":"rhwp-api v0.test","format":"hwp","warnings":["experimental convert-to-editable"]}\n' "$output"
  exit 0
fi
if [ "$cmd" = "native-op" ]; then
  output=""
  op=""
  query=""
  while [ "$#" -gt 0 ]; do
    if [ "$1" = "--output" ]; then
      shift
      output="$1"
    elif [ "$1" = "--op" ]; then
      shift
      op="$1"
    elif [ "$1" = "--query" ]; then
      shift
      query="$1"
    fi
    shift
  done
  if [ -n "$output" ]; then
    printf 'fake native op hwp' > "$output"
  fi
  if [ "$op" = "search-all-text" ]; then
    printf '{"operation":"%s","result":{"matches":[{"query":"%s","text":"before before","paragraph":0,"offset":0}]},"output":"%s","engineVersion":"rhwp-api v0.test","format":"hwp","warnings":["experimental native-op"]}\n' "$op" "$query" "$output"
    exit 0
  fi
  printf '{"operation":"%s","result":{"ok":true},"output":"%s","engineVersion":"rhwp-api v0.test","format":"hwp","warnings":["experimental native-op"]}\n' "$op" "$output"
  exit 0
fi
if [ "$cmd" = "save-as-hwp" ]; then
  output=""
  while [ "$#" -gt 0 ]; do
    if [ "$1" = "--output" ]; then
      shift
      output="$1"
    fi
    shift
  done
  printf 'fake exported hwp' > "$output"
  printf '{"saved":true,"operation":"save-as-hwp","output":"%s","outputFormat":"hwp","engineVersion":"rhwp-api v0.test","format":"hwpx","warnings":["experimental save-as-hwp"]}\n' "$output"
  exit 0
fi
echo "unexpected api command: $cmd" >&2
exit 2
""");
        if (!OperatingSystem.IsWindows())
            File.SetUnixFileMode(path,
                UnixFileMode.UserRead | UnixFileMode.UserWrite | UnixFileMode.UserExecute);
        _tempPaths.Add(path);
        return path;
    }

    private string CreateInput(string extension)
    {
        var path = Path.Combine(Path.GetTempPath(), $"bridge-input-{Guid.NewGuid():N}{extension}");
        File.WriteAllText(path, "before before");
        _tempPaths.Add(path);
        return path;
    }

    private string CreateOutput(string extension)
    {
        var path = Path.Combine(Path.GetTempPath(), $"bridge-output-{Guid.NewGuid():N}{extension}");
        _tempPaths.Add(path);
        return path;
    }

    private string CreateDirectory()
    {
        var path = Path.Combine(Path.GetTempPath(), $"bridge-out-{Guid.NewGuid():N}");
        Directory.CreateDirectory(path);
        _tempPaths.Add(path);
        return path;
    }

    private sealed record ProcessResult(int ExitCode, string Stdout, string Stderr);
}
