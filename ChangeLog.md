### [2020.07.7-alpha]
 > * Fix *sub* `NumeralRU` **исправлена ошибка подсчёта `secondDigit`**
 > * Fix *function* `GetSupplerRec` **исправлена ошибка**
 > * Fix *function* `GetRecord` **переработано**
 > * Fix *sub* `DeleteModulesAndCode` **переработано**
 * Fix *sub* `Auto_Open` **отобразить прокрутку и ярлыки листов**
 * Fix *sub* `GetForm_DialogElements` **переработана форма для `Application.Version >= 16`**
 * Fix *sub* `ChoiceCategory` **исправлено условие поиска на [ЁЕ]**
 * Fix *sub* `GetBanks` **Поиск `len(objBankID) = 5`; сортировка `objBankSUPP("Data")` по возрастанию**
 - Add *function* `Taxpayer_Number_CRC` **Проверка контрольной суммы ИНН**
 - Add *function* `PorterStemmerRU` **Стеммер Мартина Портера для русского языка**
 - Add *function* `RemoveEndings` **Стеммер: удаление окончания**
 - Add *function* `FindRegions` **Стеммер: поиск 'региона r2'**
 - Add *function* `isVowel` **гласная буква**

[2020.06.7-alpha]: ../../compare/14650a0b...master-git-2.18

### [2020.04.6-alpha] \[num-formatter\]
 - Add *object* `WordForm` as *collection* **список существительных в категории чисел**
 - Add *function* `NumberFormatterRU` **возвращает множественное количество (слово из списка)**
 - Add *function* `NumeralRU` **возвращает число прописью**

[2020.04.6-alpha]: ../../compare/9a7fac4a...num-formatter

### [2020.04.5-alpha]
 > * Rename *object* ~`DialogBox`~ to `objDialogBox`
 > * Fix *function* `GetSupplerRec` **переработано, изменён параметр `IsSelectSupp_Forced = False` (достаточно найти имя)**
 > * Fix *sub* `Shell_Sort` **переработано**
 > * Fix *function* `Trip` **переработано**
 > * Fix *function* `ClearSpacesInText` **переработано**
 > * Fix *sub* `GetWorkbooks` **переработано**
 > * Fix *function* `ChoiceCategory` **переработано**
 > * Fix *function* `GetMainPath` **переработано**
 > * Fix *function* `FileUnlocked` **переработано**
 > * Fix *sub* `GetBanks` **переработано**
 > * Fix *function* `GetRecord` **переработано**
 > * Fix *function* `GetSheetID` **переработано**
 > * Fix *sub* `DeleteModulesAndCode` **переработано**
 * Fix *function* `GetUserName` **добавлен параметр `SetUserDomain`**
 * Fix *sub* `Auto_Open` **проверка значения `CRC_HOST = 0` если книга не отчёт**

[2020.04.5-alpha]: ../../commit/9a7fac4a

# 

### [2019.10.4-alpha]
 > * Rename *sub* ~`RecLog`~ to `WriteLog`
 > * Rename && Fix *sub* ~`SettingsBankID`~ to `GetBanks` **запись списка контрагентов в коллекцию `xSUPP`**
 > * Rename *object* ~`BankID`~ to `xID`
 > * Rename *object* ~`BankSUPP`~ to `xSUPP`
 > * Fix *function* `GetQuarterNumber` **переработано, диапазон изменён (не тестировано)**
 > * Fix *function* `GetSheetID` **изменён параметр `ThisBook = True` (книга с модулем макроса)**
 > * Fix *sub* `Shell_Sort` **переработано**
 * Fix *sub* `Auto_Open` **добавлен массив директорий `Paths` для отчётов (задаётся `CRC_HOST`)**
 * Fix *sub* `GetWorkbooks` **остановить макросы если книга не найдена или невозможно открыть**
 - Add *function* `Trip` **удаляет переносы строки (на краю строки)**
 - Add *function* `Tripp` **удаляет разрывы строки (на краях в массиве строк)**
 - Add *object* `DialogBox` **диалоговый лист**
 - Add *sub* `DialogButtons_Click` **события кнопок диалогового листа**
 - Add *sub* `GetForm_DialogElements` **создание диалогового листа и элементов**
 - Add *const* `CRC_HOST` **выбор директорий для отчётов (X mod 2^N)**
 - Add *function* `GetSupplerRec` **возвращает запись контрагента по дате из `xSUPP`**
 - Add *function* `GetRecord` **возвращает значение из коллекции `xID` по ключу**
 > -
 >     * Fix *function* `ClearSpacesInText` **помарки**
 >     * Fix *function* `GetUserName` **помарки**
 >     * Fix *function* `ChoiceCategory` **помарки**
 >     * Fix *function* `GetMainPath` **помарки**
 >     * Fix *function* `FileUnlocked` **помарки**
 >     * Fix *sub* `DeleteModulesAndCode` **помарки**

[2019.10.4-alpha]: ../../commit/96cc161f

# 

### [2018.08.3-alpha]
 > * Rename *function* ~`FindSheet`~ and parameters to `GetSheetID`
 > * Rename *sub* ~`Record_Log`~ to `RecLog`

[2018.08.3-alpha]: ../../commit/9f422069

### [2018.07.2-alpha]
 > * Rename *function* ~`GetSheetIndex`~ and parameters to `FindSheet`
 > * Fix *function* `ClearSpacesInText` **переработано**
 > * Fix *sub* `Auto_Open` **подготовка данных о книгах**
 - Add *sub* `Shell_Sort` **сортировка методом Шелла Дональда**
 - Add *function* `GetQuarterNumber` **возвращает номер квартала**
 - Add *property* `GetUserName` **возвращает имя текущей учётной записи**
 - Add *sub* `SettingsBankID` **структура книги (именованные диапазоны)**
 - Add *function* `ChoiceCategory` **возвращает индекс категории контрагента**
 - Add *property* `GetMainPath` **возвращает полный путь к книгам**
 - Add *sub* `Record_Log` **журнал открытия книг**
 - Add *function* `FileUnlocked` **возвращает текущий статус файла**
 - Add *sub* `DeleteModulesAndCode` **удаляет модули из книги**
 - Add *object* `DirName` as *collection* **список полных путей к файлам**
 - Add *object* `FileName` as *collection* **список имён файлов с расширением**
 - Add *object* `Manager` as *collection* **список ответственных менеджеров файлов**
 - Add *object* `BankID` as *collection*
 - Add *object* `BankSUPP` as *collection*
 - Add *sub* `GetWorkbooks` **запись найденных книг в коллекцию `DirName`**

[2018.07.2-alpha]: ../../commit/9a2087e4

# 

### [2015.10.1-alpha] \[dev-heavy-old\]
 > - Add *object* `ARCH_` as *worksheet*
 > - Add *sub* `Auto_Open` **после открытия возвращать статус "Сохранено", сортировка цен**
 > - Add *sub* `SortSupplier` **сортировка данных листа**
 - Add *function* `ClearSpacesInText` **удаляет лишние пробелы и непечатаемые символы**
 - Add *function* `GetSheetIndex` **возвращает индекс листа**

[2015.10.1-alpha]: ../../../StatsOKM/compare/e784ad25...dev

# 
