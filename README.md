# Docx4J Open XML Fluent API

[![Maven Central](https://img.shields.io/maven-central/v/io.github.sfali23/docx4j-builder.svg?label=Maven%20Central)](https://search.maven.org/search?q=g:%22io.github.sfali23%22%20AND%20a:%22docx4j-builder%22)
[![License](https://img.shields.io/badge/License-Apache%202.0-blue.svg)](https://opensource.org/licenses/Apache-2.0)
[![Java Version](https://img.shields.io/badge/Java-21-blue.svg)](https://www.oracle.com/java/technologies/javase/jdk21-archive-downloads.html)

A fluent API for building Open XML documents using Docx4J. This library provides a type-safe, intuitive interface for creating and manipulating Word documents programmatically.

## Features

- 🚀 **Fluent API**: Intuitive, chainable methods for document creation
- 📝 **Type-Safe**: Compile-time safety with strong typing
- 🎨 **Rich Formatting**: Support for styles, tables, paragraphs, and more
- 🔧 **Built on Docx4J**: Leverages the power of the Docx4J library
- ✅ **Well-Tested**: Comprehensive test coverage
- 📦 **Maven Central**: Easy integration via Maven/Gradle

## Installation

### Gradle (Kotlin DSL)

```kotlin
dependencies {
    implementation("io.github.sfali23:docx4j-builder:0.5.6")
}
```

### Gradle (Groovy DSL)

```groovy
dependencies {
    implementation 'io.github.sfali23:docx4j-builder:0.5.6'
}
```

### Maven

```xml
<dependency>
    <groupId>io.github.sfali23</groupId>
    <artifactId>docx4j-builder</artifactId>
    <version>0.5.6</version>
</dependency>
```

## Requirements

- Java 21 or higher
- Docx4J 11.4.9

## Quick Start

```java
// Example usage will be added here
// Create a document with fluent API
```

## Building from Source

### Prerequisites

- JDK 21 or higher
- Gradle 9.4.1 or higher (wrapper included)

### Build Commands

```bash
# Build the project
make build
# or
./gradlew build

# Run tests
make test
# or
./gradlew test

# Clean build
make clean
# or
./gradlew clean

# Publish to local Maven repository
make publishLocal
# or
./gradlew publishToMavenLocal
```

## Project Structure

```
open-xml-builder/
├── app/                    # Main library module (docx4j-builder)
├── buildSrc/              # Custom Gradle plugins
├── scripts/               # Build and publishing scripts
└── build.gradle.kts      # Root build configuration
```

## Dependencies

- **Docx4J Core**: 11.4.9
- **Docx4J Export FO**: 11.4.9
- **SLF4J API**: 2.0.11
- **Logback Classic**: 1.4.14
- **AlphaSystem Commons**: 0.3.2

## Contributing

Contributions are welcome! Please follow these guidelines:

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Use conventional commits for your messages
4. Commit your changes (`git commit -m 'feat: add amazing feature'`)
5. Push to the branch (`git push origin feature/amazing-feature`)
6. Open a Pull Request

## License

This project is licensed under the Apache License 2.0 - see the [LICENSE.md](LICENSE.md) file for details.

## Links

- **GitHub Repository**: https://github.com/AlphaSystemSolution/open-xml-builder
- **Maven Central**: https://search.maven.org/artifact/io.github.sfali23/docx4j-builder
- **Issues**: https://github.com/AlphaSystemSolution/open-xml-builder/issues

## Acknowledgments

- Built with [Docx4J](https://www.docx4java.org/)

## Version History

See [Releases](https://github.com/AlphaSystemSolution/open-xml-builder/releases) for version history and changelogs.
