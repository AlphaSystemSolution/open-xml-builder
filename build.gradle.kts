plugins {
    `maven-publish`
    id("io.codearte.nexus-staging") version "0.30.0"
}

allprojects {
    repositories {
        mavenLocal()
        mavenCentral()
        maven {
            url = uri("https://s01.oss.sonatype.org/content/repositories/releases/")
        }
    }
}

// Configure Nexus Staging plugin
nexusStaging {
    username = project.findProperty("ossrhUsername") as String? ?: System.getenv("OSSRH_USERNAME")
    password = project.findProperty("ossrhPassword") as String? ?: System.getenv("OSSRH_PASSWORD")
    packageGroup = "io.github.sfali23"
    stagingProfileId = project.findProperty("stagingProfileId") as String? ?: System.getenv("STAGING_PROFILE_ID")
    
    // Configure staging repository actions
    numberOfRetries = 10
    delayBetweenRetriesInMillis = 5000
}
