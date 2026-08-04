namespace OfficeCli.Tests;

public sealed class InstallerSecurityTests
{
    [Fact]
    public void WindowsInstaller_RequiresVersionedChecksummedSidecars()
    {
        var script = File.ReadAllText(Path.Combine(FindRepoRoot(), "install.ps1"));

        Assert.Contains("$version -and $sidecarChecksumsAvailable", script);
        Assert.Contains("Get-FileHash -Path $sidecarTemp -Algorithm SHA256", script);
        Assert.Contains("$parts[1] -eq $sidecarAsset", script);
        Assert.Contains("refusing mutable remote sidecar URL", script);
        Assert.DoesNotContain("releases/latest/download/$sidecarAsset", script);
    }

    [Fact]
    public void UnixInstaller_RequiresVersionedChecksummedSidecars()
    {
        var script = File.ReadAllText(Path.Combine(FindRepoRoot(), "install.sh"));

        Assert.Contains("[ -n \"$VERSION\" ]", script);
        Assert.Contains("SIDECAR_CHECKSUMS_AVAILABLE", script);
        Assert.Contains("awk -v a=\"$sidecar_asset\" '$2 == a", script);
        Assert.Contains("refusing mutable remote sidecar URL", script);
    }

    private static string FindRepoRoot()
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory != null)
        {
            if (File.Exists(Path.Combine(directory.FullName, "install.ps1"))
                && File.Exists(Path.Combine(directory.FullName, "install.sh")))
                return directory.FullName;
            directory = directory.Parent;
        }
        throw new DirectoryNotFoundException("OfficeCLI repository root not found.");
    }
}
