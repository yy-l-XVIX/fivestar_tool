build-win-w:
	@pyinstaller -F -a \
		-n main_win_v1 \
		--clean \
		--add-data "form_template.docx;form_template.docx" \
		--distpath dist \
		main.py

build-mac-m1-w:
	@pyinstaller -F -a \
		-n main_mac_m1_v1 \
		--clean \
		--add-data form_template.docx:form_template.docx \
		--distpath dist \
		main.py

zip-for-mac:
	@zip main_mac_m1_v1.zip dist/main_mac_m1_v1