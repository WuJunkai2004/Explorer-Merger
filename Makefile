all:
	python scripts/build.py

debug:
	python scripts/compile.py

clean:
	rm -rf dist
	rm *.obj
	rm *.pdb

.PHONY: all debug clean
