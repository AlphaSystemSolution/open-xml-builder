plugins {
    `maven-publish`
    signing
    id("net.researchgate.release") version "3.1.0"
    id("io.github.gradle-nexus.publish-plugin") version "2.0.0"
}

apply(from = "${rootDir}/scripts/nexus-publish.gradle")

allprojects {
    repositories {
        mavenLocal()
        mavenCentral()
        maven {
            url = uri("https://s01.oss.sonatype.org/content/repositories/releases/")
        }
    }
}

configure<net.researchgate.release.ReleaseExtension> {
    tagTemplate.set("v\${version}")
}

afterEvaluate {
    tasks.named("afterReleaseBuild") {
        dependsOn(
            tasks.named("initializeSonatypeStagingRepository"),
            tasks.named("publishToSonatype"),
            tasks.named("closeAndReleaseSonatypeStagingRepository")
        )
    }
}
