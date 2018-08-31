## Известные проблемы при работе с репозиторием

### Проблемы чтения кодировки windows-1251 в репозитории

1. При просмотре в [репозитории] по умолчанию файлы отображаются в кодировке `UTF-8`.
2. При редактировании в [репозитории] файлы пересохраняются в кодировке `UTF-8`.
3. Отредактированные файлы в папке `/src/*.*` после [скачивания в ZIP] 
необходимо открыть в **Блокноте** и пересохранить как `ANSI`. При импорте 
модули должны быть в кодировке `windows-1251`.
4. Дополнительно: после импортирования модулей зайти в "Параметры" **Excel** -> 
"Центр управления безопасностью", в меню "Параметры конфиденциальности" 
отключить параметр "Удалять персональные данные из свойств файла при сохранении".

[репозитории]://github.com/bopoh13/ReptOKM/tree/master/src
[скачивания в ZIP]://github.com/bopoh13/ReptOKM/archive/master.zip

### Настройка кодировки файлов в Windows через фильтр Git

1. В корне (клона) репозитория необходимо создать файл `.gitattributes` и указать 
в нём файлы
	``` markdown
	# Custom for Visual Basic (CRLF for classes or modules)
	*.bas	filter=win1251  eol=crlf
	*.cls	filter=win1251  eol=crlf
	```

2. Задать фильтр для файлов и отключить замену окончаний строк
	``` bash
	$ git config --global filter.win1251.clean "iconv -f windows-1251 -t utf-8"
	$ git config --global filter.win1251.smudge "iconv -f utf-8 -t windows-1251"
	$ git config --global filter.win1251.required true
	# Не изменять окончания строк в репозитории
	$ git config core.autocrlf false
	$ git config core.eol crlf
	```
3. Проект, редактируемый стандартными средствами Windows, имеет символы возврата 
каретки, но они удаляются из файлов в кодировке `UTF-8` вместе с BOM. Для 
последующего добавления символа возврата каретки используйте команду `git cr`
	``` bash
	$ git config --global alias.cr\
	 '!find . -type f \( -name "*.md" -o -name "*.xml" \) -print0\
	 | xargs -0 grep -m1 -l `printf "^\xEF\xBB\xBF"`\
	 | xargs sed -i "1 s/^\xEF\xBB\xBF//; s/\$/\x0D/"\
	 && git ls-files -mo --eol'
	```

4. Установить флаг `--no-ff`, чтобы Git всегда создавал новый объект Commit при 
слиянии. Информация о существующей ветке не потеряется.
	``` bash
	$ git config --global merge.ff false
	```

5. Теперь можно работать с файлами через Git Bash или Git Client не заботясь 
о кодировке.

# 
