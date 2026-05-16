GRADLE = ./gradlew

build:
	$(GRADLE) build --refresh-dependencies

clean:
	$(GRADLE) clean

test:
	$(GRADLE) test

all: clean build test

publishLocal:
	$(GRADLE) publishToMavenLocal

release:
	$(GRADLE) publishAndReleaseToMavenCentral
