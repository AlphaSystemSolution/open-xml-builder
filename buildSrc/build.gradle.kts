repositories {
    mavenLocal()
    mavenCentral()
    maven {
        url = uri("https://repo.maven.apache.org/maven2")
    }
}

val docx4jVersion: String by project

dependencies {
    implementation("com.sun.codemodel:codemodel:2.6")
    implementation("org.docx4j:docx4j-core:11.5.13")
    implementation("org.apache.commons:commons-text:1.15.0")
}
