GRADLE = ./gradlew

.PHONY: build clean, test all spotless publishLocal release

build:
	$(GRADLE) build

clean:
	$(GRADLE) clean

test:
	$(GRADLE) test

all: clean build test

spotless:
	$(GRADLE) spotlessApply

publishLocal:
	$(GRADLE) publishToMavenLocal

release:
	$(GRADLE) publishToMavenCentral createTag pushTag
