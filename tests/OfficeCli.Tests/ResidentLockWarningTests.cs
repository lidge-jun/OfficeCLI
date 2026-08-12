// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

namespace OfficeCli.Tests;

public class ResidentLockWarningTests
{
    [Fact]
    public void CreateResidentSuffixNamesLockConsequencesAndCloseRecovery()
    {
        var suffix = CommandBuilder.FormatCreatedResidentSuffix("lock.xlsx");

        Assert.Contains("kept open by a background resident", suffix);
        Assert.Contains("may remain locked", suffix);
        Assert.Contains("officecli close \"lock.xlsx\"", suffix);
        Assert.Contains("moving, renaming, deleting", suffix);
        Assert.Contains("opening it in another program", suffix);
        Assert.DoesNotContain("faster subsequent commands", suffix);
    }
}
