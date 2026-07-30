import com.alphasystem.openxml.gradleplugin.CodeGenerator

plugins {
    `java-library`
    alias(libs.plugins.publish)
    alias(libs.plugins.semver.release)
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

val isLocalPublish = gradle.startParameter.taskNames.any { it.contains("publishToMavenLocal") }

mavenPublishing {
    publishToMavenCentral(automaticRelease = true)
    if (!isLocalPublish) {
        signAllPublications()
    }

    coordinates("io.github.sfali23", "docx4j-builder")

    pom {
        name.set("Docx4J Builder")
        description.set("Docx4J Open XML Fluent API")
        url.set("https://github.com/AlphaSystemSolution/open-xml-builder")

        licenses {
            license {
                name.set("The Apache License, Version 2.0")
                url.set("http://www.apache.org/licenses/LICENSE-2.0.txt")
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

tasks.withType<Test>().configureEach {
    systemProperty("docs.dir", "build/docs")
}

tasks.named<Test>("test") {
    useTestNG {
        suites("testng/testng.xml")
    }
}

semverrelease {
    addUnReleasedCommitsToTagComment.set(true)
}

