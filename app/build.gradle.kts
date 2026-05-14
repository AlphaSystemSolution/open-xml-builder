import com.alphasystem.openxml.gradleplugin.CodeGenerator

plugins {
    `java-library`
    id("com.vanniktech.maven.publish") version "0.36.0"
}

repositories {
    mavenLocal()
    mavenCentral()
}

val asCommonsVersion: String by project
val docx4jVersion: String by project
val moxyVersion: String by project
val eclipseMoxyVersion: String by project
val slf4jApiVersion: String by project
val logbackClassicVersion: String by project
val testngVersion: String by project
val guiceVersion: String by project
val reportngVersion: String by project
val commonsCodecVersion: String by project
val batikVersion: String by project
val guavaVersion: String by project

dependencies {
    api("io.github.sfali23:commons:$asCommonsVersion")
    api("org.docx4j:docx4j-core:$docx4jVersion")
    api("org.docx4j:docx4j-export-fo:$docx4jVersion")
    api("org.docx4j:docx4j-JAXB-MOXy:$docx4jVersion")
    api("org.docx4j:docx4j-MOXy-JAXBContext:$moxyVersion")
    api("org.eclipse.persistence:org.eclipse.persistence.moxy:$eclipseMoxyVersion")
    api("org.docx4j:docx4j-JAXB-ReferenceImpl:$docx4jVersion")
    api("org.slf4j:slf4j-api:$slf4jApiVersion")
    api("ch.qos.logback:logback-classic:$logbackClassicVersion")
    testImplementation("org.testng:testng:$testngVersion")
    testImplementation("com.google.inject:guice:$guiceVersion")
    testImplementation("org.uncommons:reportng:$reportngVersion")
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

    coordinates("io.github.sfali23", "docx4j-builder", "$version")

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

