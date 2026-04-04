# Release Process Fix - April 3, 2026

## Issue Fixed
The `make release` command was failing with the error:
```
No staging repository with name sonatype created
```

## Root Cause
The `afterReleaseBuild` task was trying to close and release the Sonatype staging repository before it was properly initialized.

## Solution Applied
Updated `build.gradle.kts` to include the initialization task in the dependency chain:

```kotlin
afterEvaluate {
    tasks.named("afterReleaseBuild") {
        dependsOn(
            tasks.named("initializeSonatypeStagingRepository"),
            tasks.named("publishToSonatype"),
            tasks.named("closeAndReleaseSonatypeStagingRepository")
        )
    }
}
```

## Task Execution Order
The release process now executes tasks in the correct order:
1. `initializeSonatypeStagingRepository` - Creates/finds the staging repository
2. `publishToSonatype` - Publishes artifacts to the staging repository
3. `closeAndReleaseSonatypeStagingRepository` - Closes and releases the repository

## Testing Performed
- ✅ Dry run completed successfully
- ✅ Task dependencies verified
- ✅ Publishing to local Maven repository works
- ✅ Credentials file exists at `~/.gradle/sonatype.properties`

## Next Steps
To perform a release:

### Option 1: Automatic Release (Non-Interactive)
```bash
make release
# or
./gradlew release -Prelease.useAutomaticVersion=true
```

### Option 2: Interactive Release
```bash
./gradlew release
```
This will prompt you for:
- Release version (e.g., 0.5.6)
- Next development version (e.g., 0.5.7-SNAPSHOT)

### Option 3: Test Without Publishing
To test the release process without publishing to Sonatype:
```bash
./gradlew release -Prelease.useAutomaticVersion=true -x publishToSonatype -x closeAndReleaseSonatypeStagingRepository
```

## Post-Release Verification
After a successful release:
1. Verify the Git tag was created: `git tag -l`
2. Push the tag: `git push origin v<version>`
3. Push commits: `git push`
4. Check Sonatype: https://s01.oss.sonatype.org/
5. Wait for Maven Central sync (2-4 hours)

## Current Version
- Current: `0.5.6-SNAPSHOT`
- Last released: `v0.5.5`

## Files Modified
- `build.gradle.kts` - Fixed task dependencies in `afterReleaseBuild`
