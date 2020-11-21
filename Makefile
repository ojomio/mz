
PHONY: gui

gui: qt/*.py

%.py: %.ui
	pyuic5 $< > $@
