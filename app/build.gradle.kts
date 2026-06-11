import com.alphasystem.openxml.gradleplugin.CodeGenerator
import net.researchgate.release.ReleaseExtension

plugins {
    `java-library`
    signing
    `maven-publish`
    id("net.researchgate.release") version "3.1.0"
}

repositories {
    mavenLocal()
    mavenCentral()
}

dependencies {
    api(libs.alphasystm.commons)
    api(libs.docx4j.core)
    api(libs.docx4j.exportfo)
    api(libs.docx4j.jaxbmoxy)
    api(libs.docx4j.jaxbreferenceimpl)
    api(libs.docx4j.moxyjaxbcontext)
    api(libs.docx4j.eclipsemoxy)
    api(libs.slf4j.api)
    api(libs.logback.classic)
    testImplementation(libs.testng)
    testImplementation(libs.guice)
    testImplementation(libs.guava)
    testImplementation(libs.reportng)
}

group = "io.github.sfali23"

val generatedSrcDir = file("${layout.buildDirectory.get()}/generated/src/main/java")

sourceSets {
    main {
        java {
            srcDir(generatedSrcDir)
        }
    }
}

tasks.register("generateCode") {
    doLast {
        CodeGenerator.wmlGenerator(generatedSrcDir)
    }
}

tasks.named<JavaCompile>("compileJava") {
    dependsOn("generateCode")
}

tasks.withType<JavaCompile>().configureEach {
    options.encoding = "UTF-8"
}

java {
    sourceCompatibility = JavaVersion.VERSION_21
    targetCompatibility = JavaVersion.VERSION_21
    withSourcesJar()
}

configure<ReleaseExtension> {
    ignoredSnapshotDependencies.set(listOf("net.researchgate:gradle-release"))
    with(git) {
        requireBranch.set("main")
        pushToRemote.set("origin")
    }
    tagTemplate.set("v\$version")
    preTagCommitMessage.set("Pre tag commit: ")
    tagCommitMessage.set("Release version ")
    newVersionCommitMessage.set("Next version development")
}

// Configure release plugin to publish to Maven Central and push GitHub tags
tasks.named("release") {
    dependsOn("signArchives", "publishToSonatype", "closeAndReleaseSonatypeStagingRepository")
}

tasks.named("afterReleaseBuild") {
    doLast {
        println("Release completed successfully!")
        println("Published to Maven Central and pushed tag to GitHub")
    }
}

// Custom task to publish to Sonatype (Maven Central)
tasks.register("publishToSonatype") {
    dependsOn("publishMavenPublicationToOSSRHRepository")
}


publishing {
    publications {
        create<MavenPublication>("maven") {
            from(components["java"])
            
            groupId = "io.github.sfali23"
            artifactId = "docx4j-builder"
            version = project.version.toString()
            
            pom {
                name.set("Docx4J Builder")
                description.set("Docx4J Open XML Fluent API")
                url.set("https://github.com/AlphaSystemSolution/open-xml-builder")
                
                licenses {
                    license {
                        name.set("Apache-2.0")
                        url.set("https://www.apache.org/licenses/LICENSE-2.0")
                    }
                }
                
                developers {
                    developer {
                        id.set("sfali23")
                        name.set("Syed Farhan Ali")
                        email.set("f.syed.ali@gmail.com")
                    }
                }
                
                scm {
                    connection.set("scm:git:git://github.com/AlphaSystemSolution/open-xml-builder.git")
                    developerConnection.set("scm:git:ssh://github.com/AlphaSystemSolution/open-xml-builder.git")
                    url.set("https://github.com/AlphaSystemSolution/open-xml-builder")
                }
            }
        }
    }
    
    repositories {
        maven {
            name = "MavenCentral"
            url = uri("https://ossrh-staging-api.central.sonatype.com/service/local/")
            credentials {
                username = project.findProperty("ossrhUsername") as String? ?: System.getenv("OSSRH_USERNAME")
                password = project.findProperty("ossrhPassword") as String? ?: System.getenv("OSSRH_PASSWORD")
            }
        }
    }
}

// Configure jar signing
signing {
    sign(publishing.publications["maven"])
}

// Configure signing credentials
signing {
    useInMemoryPgpKeys(
        project.findProperty("signing.secretKey") as String? ?: System.getenv("SIGNING_SECRET_KEY"),
        project.findProperty("signing.password") as String? ?: System.getenv("SIGNING_PASSWORD")
    )
}

// Only sign when not building a snapshot version
tasks.withType<Sign>().configureEach {
    onlyIf { !project.version.toString().endsWith("-SNAPSHOT") }
}

tasks.withType<Test>().configureEach {
    systemProperty("docs.dir", "build/docs")
}

tasks.named<Test>("test") {
    useTestNG {
        suites("testng/testng.xml")
    }
}

