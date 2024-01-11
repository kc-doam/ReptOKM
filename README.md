## Известные проблемы при работе с удалённым хранилищем

<a name="setup-repo-after-copy-files"></a>
### Настройка репозитория после копирования файлов

1. При просмотре в [удалённом хранилище] по умолчанию файлы отображаются в кодировке `UTF-8`.

2. При редактировании в [удалённом хранилище] файлы пересохраняются в кодировке `UTF-8`.

3. Отредактированные файлы в папке `<REPO>/src/*.*` после [скачивания в ZIP] 
   необходимо открыть в **Текстовом редакторе** в кодировке `windows-1251`.

4. Чтобы после импортирования модулей каждый раз при сохранении (и автосохранении) 
   не отображалось сообщение "*Будьте внимательны! В документе могут быть персональные 
   данные, которые невозможно удалить с помощью инспектора документов.*" 
   нужно зайти в "Параметры" **Excel** -> "Центр управления безопасностью", 
   в меню "Параметры конфиденциальности" отключить параметр 
   "Удалять персональные данные из свойств файла при сохранении".  
   Тоже самое сделает макрос:
   ``` vba
   Sub ConfidentialInformationAlert_Disable()
     If ActiveWorkbook.RemovePersonalInformation Then ActiveWorkbook.RemovePersonalInformation = False
   End Sub
   ```

[удалённом хранилище]: ../master/src
[скачивания в ZIP]: ../../archive/master.zip

<a name="set-codepage-1251"></a>
### Настройка кодировки файлов windows-1251 для модулей через фильтр Git

1. Чтобы Git старых версий сохранял изменения файлов без BOM автоматически 
   нужно скачать [клиентский хук] в директорию `<REPO>/.git/hooks` без расширения `.sh`  
   Хук будет выполняться перед сохранением изменений (*commit*).

2. Установить флаг `--no-ff`, чтобы Git всегда создавал отдельный объект с изменениями 
   перед слиянием. Информация о существующей ветви не потеряется.
   ``` console
   $ git config --local merge.ff false
   ```

3. Используйте текстовый редактор, который при сохранении файлов оставляет только 
   единственный символ переноса строк `crlf`. Репозиторий имеет настройки для работы в 
   Git Bash или Git Desktop Client. Для работы в текстовом редакторе **vscode** 
   необходимо добавить следующие настройки:
   ``` json
   "[vb][vba]": {
   	"editor.defaultFormatter": "serkonda7.vscode-vba",
   	"editor.insertSpaces": true,
   	"editor.fontFamily": "Menlo, Monaco, 'Courier New', monospaces",
   	"editor.fontSize": 13,
   	"editor.language.brackets": [ [ "(", ")" ] ],
   	"editor.lineHeight": 16,
   	"editor.maxTokenizationLineLength": 2000,
   	"editor.rulers": [ 80, { "column": 120, "color": "#ff0000" } ],
   	"editor.tabSize": 2,
   	"files.autoGuessEncoding": true,
   	"files.encoding": "windows1251",
   	"files.eol": "\r\n",
   	"files.insertFinalNewline": true
   },
   ```

[клиентский хук]: https://gist.github.com/c55f1538454755fdff71fba0d686e371

<details>
    <summary><a name="shell-sort-gap"><picture><source media="(prefers-color-scheme: dark)" srcset="https://cdn.simpleicons.org/pkgsrc/fff?raw=true" type="image/svg+xml"><img src="https://cdn.simpleicons.org/pkgsrc/000?raw=true" type="image/svg+xml" alt="cube" align="left" width="24" height="24"/></picture></a> Интервалы для сортировки методом Шелла Дональда</summary><br />
  
   |OEIS|Name Gap|Complexity[^1]|Formula|[&fnof;(k)]
   |:------- | ---:|:---:|:--- |:--- 
   |[A102549]|Ciura 2001|Un­known| |`={ 1750; 701; 301; 132; 57; 23; 10; 4; 1 }`
   |[A108870]|Tokuda 1992|Un­known| |`=ОКРУГЛВВЕРХ(( 9*(9/4)^A4 -4 )/5;0)`
   |[A033622]|Sedgewick 1986|$\theta( N^\frac{4}{3} )$|`=ЕСЛИ( ЕНЕЧЁТ(k); 8*2^k -6*2^( (k+1)/2 ); 9*2^k -9*2^(k/2) ) +1`|`=( 9-ОСТАТ(k;2) )*2^k -( 9-3*ОСТАТ(k;2) )*2^ОКРУГЛВВЕРХ(k/2;0) +1`
   |[A055875]|Knuth 1973|$\theta( N^\frac{3}{2} )$| |`=( 3^k -1 )/2`
   |[A003586]|Pratt 1971|$\theta( N \times lg^2 (N) )$|$\sum \limits_{k=1}^{N/_2} 2^k \times 3^k$|` `
   |[A033547]|Shell 1959|$\theta( N^2 )$| |`=ОКРУГЛВНИЗ(N/2^k;0)`

   > :warning: 
   > Последовательности со степенями числа 2 уменьшают эффективность сортировки.  

   [^1]: **Complexity** - Worst-case time complexity.
</details>

[Shellsort]: https://en.wikipedia.org/wiki/Shellsort
[сортировка]: https://neerc.ifmo.ru/wiki/index.php?title=Сортировки
[&fnof;(k)]: ../../search?q=Shell_Sort&type=code
[A102549]: https://oeis.org/A102549
[A108870]: https://oeis.org/A108870
[A033622]: https://oeis.org/A033622
[A055875]: https://oeis.org/A055875
[A003586]: https://oeis.org/A003586
[A033547]: https://oeis.org/A033547

