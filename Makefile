all:
	python scripts/build.py

debug:
	python scripts/compile.py

.PHONY: all debug
