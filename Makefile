
PHONY: gui

gui: print/src/main/python/qt/*.py

%.py: %.ui
	pyuic5 $< > $@

CODE=.

lint:
	black --target-version py38 --check --skip-string-normalization --line-length=100 $(CODE)
	flake8 --statistics $(CODE)
	pylint --rcfile=setup.cfg $(CODE)

pretty:
	black --target-version py38 --skip-string-normalization --line-length=100 $(CODE)
	isort  $(CODE)

build-installer:
	cd print && fbs freeze && fbs installer
