plugins {
    `maven-publish`
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
