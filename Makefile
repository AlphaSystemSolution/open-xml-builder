GRADLE = ./gradlew

build:
	$(GRADLE) build

clean:
	$(GRADLE) clean

test:
	$(GRADLE) test

all: clean build test

publishLocal:
	$(GRADLE) setReleaseVersion publishToMavenLocal

release:
	$(GRADLE) setReleaseVersion publishToMavenCentral createTag pushTag
