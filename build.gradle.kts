plugins {
    id("java")
}

group = "org.example"
version = "1.0-SNAPSHOT"

repositories {
    mavenCentral()
}

dependencies {
    testImplementation(platform("org.junit:junit-bom:5.9.1"))
    testImplementation("org.junit.jupiter:junit-jupiter")
    implementation(group = "org.apache.poi", name = "poi", version = "5.0.0")
    implementation(group = "org.apache.poi", name = "poi-ooxml", version = "5.0.0")
}


tasks.test {
    useJUnitPlatform()
}
