repositories {
    mavenLocal()
    mavenCentral()
    maven {
        url = uri("https://repo.maven.apache.org/maven2")
    }
}

dependencies {
    implementation(libs.codemodel)
    implementation(libs.docx4j.core)
    implementation(libs.commons.text)
}
