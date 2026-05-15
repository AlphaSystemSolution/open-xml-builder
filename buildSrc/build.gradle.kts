repositories {
    mavenLocal()
    mavenCentral()
    maven {
        url = uri("https://repo.maven.apache.org/maven2")
    }
}

dependencies {
    implementation("com.sun.codemodel:codemodel:${libs.versions.codemodel.get()}")
    implementation("org.docx4j:docx4j-core:${libs.versions.docx4j.get()}")
    implementation("org.apache.commons:commons-text:${libs.versions.commonsText.get()}")
}
