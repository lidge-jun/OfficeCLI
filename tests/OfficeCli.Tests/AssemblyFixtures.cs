// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

// The HWP test suite drives the CLI in-process: InvokeOfficeCli redirects
// Console.Out to a StringWriter to capture output, and the bridge tests set
// process-wide OFFICECLI_RHWP_* variables to point at fake sidecars.
//
// Both are global process state. xunit's [Collection] only serializes tests
// within a collection -- separate collections still run concurrently by
// default, so a test in one collection could steal another's redirected stdout
// mid-invocation. That produced a genuinely nondeterministic suite: the same
// commit yielded 0, then 6, then 6 failures across consecutive runs, always in
// tests that assert on captured stdout.
//
// Disabling assembly-level parallelization is the honest fix. The alternative
// -- rewriting every test to spawn a real subprocess -- is a larger change to
// restored fork code than this port should carry, and the suite is ~2 minutes
// serial.
[assembly: Xunit.CollectionBehavior(DisableTestParallelization = true)]
