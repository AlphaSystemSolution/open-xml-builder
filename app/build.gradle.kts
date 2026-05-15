import com.alphasystem.openxml.gradleplugin.CodeGenerator

plugins {
    `java-library`
    id("com.vanniktech.maven.publish") version "0.36.0"
}

repositories {
    mavenLocal()
    mavenCentral()
}

dependencies {
    api("io.github.sfali23:commons:${libs.versions.asCommons.get()}")
    api("org.docx4j:docx4j-core:${libs.versions.docx4j.get()}")
    api("org.docx4j:docx4j-export-fo:${libs.versions.docx4j.get()}")
    api("org.docx4j:docx4j-JAXB-MOXy:${libs.versions.docx4j.get()}")
    api("org.docx4j:docx4j-JAXB-ReferenceImpl:${libs.versions.docx4j.get()}")
    api("org.docx4j:docx4j-MOXy-JAXBContext:${libs.versions.moxy.get()}")
    api("org.eclipse.persistence:org.eclipse.persistence.moxy:${libs.versions.eclipseMoxy.get()}")
    api("org.slf4j:slf4j-api:${libs.versions.slf4jApi.get()}")
    api("ch.qos.logback:logback-classic:${libs.versions.logbackClassic.get()}")
    testImplementation("org.testng:testng:${libs.versions.testng.get()}")
    testImplementation("com.google.inject:guice:${libs.versions.guice.get()}")
    testImplementation("com.google.guava:guava:${libs.versions.guava.get()}")
    testImplementation("org.uncommons:reportng:${libs.versions.reportng.get()}")
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

mavenPublishing {
    publishToMavenCentral(automaticRelease = true)
    signAllPublications()

    coordinates("io.github.sfali23", "docx4j-builder", libs.versions.libVersion.get())

    pom {
        name.set("Open XML Builder")
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

tasks.withType<Test>().configureEach {
    systemProperty("docs.dir", "build/docs")
}

tasks.named<Test>("test") {
    useTestNG {
        suites("testng/testng.xml")
    }
}

