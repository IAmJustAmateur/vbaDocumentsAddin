# vbaDocumentsAddin
VBA macros for automating the filling out of contracts, work orders, invoices and various kinds of documents with tables
Надстройка vba для автоматизации заполнения договоров и прочих документов, таких как заказ-наряды на выполнение работ, акты выполненных работ и счета-фактуры
## Описание файлов
### форма
OrderForm.frm - форма заказ-наряда, код
OrderForm.frx - форма заказ-наряда, внутренне представление формы
### модули
errors.bas - исключения, сообщения об ошибках
factories.bas - фабрика класса СRowFrame, строка таблицы формы заказ-наряда
NamesInGenitive.bas - ФИО в родительном падеже, использование - см. в модуле tests.bas: test_name_in_genitive()
tests.bas - тесты, примеры использования
utis.bas - утилиты
### модули классов
CDateNumberBox.cls - реализация TextBox'ов для даты - день, месяц, год. OrderForm должна содержать textboxes tbDay, tbMonth, tbYear соответственно
CDateSpinButtons.cls - реализация SpinButtons для даты. OrderForm должна содержать spinbuttons: sbDay, sbMonth, sbYear
CDefaultDocument.cls - дефолтная имплементация автоматизации заполнения документа .docx: заполняются поля и табличные поля
CDefaultNameGenerator.cls - дефолтная генерация имени документа: среди полей выбирается первое поле, включающее "номер" и этот номер используется в качестве имени документа
CDefaultTables.cls - дефолтная реализация выбора таблиц документа, содержащих табличные поля, выбираются таблицы соответствующие шаблону "order_template.docx"
CFieldsReaderFromXL.cls - reader переменных полей документа из .xlsx файла
CNumberBox.cls -реализация TextBox для ввода только целых чисел с ограничением на минимум и максимум
COrder.cls - заказ-наряд
COrderFieldsReader.cls - reader полей заказ-наряда из формы OrderForm
COrderForm.cls - класс для загрузки начальных значений в Combobox'ы OrderForm - Заказчики и Работы, а также для привязки генератора документов к заказ-наряду
CPriceTextBox.cls - реализация textbox для ввода цен - число с 2мя знаками после запятой
CQtySpinButton.cls - SpinButton для количества (работ, запчастей)
CRemoveButton.cls - реализация CommandButton для удаления строки таблицы заказ-наряда (СRowFrame)
CRowFrame.cls - реализация строки заказ-наряда как MSForms.Frame
CTestCustomer.cls - тестовый класс для заполнения значений combobox cmbCustomers в OrderForm
CTestLoaders.cls - тестовый класс для загрузки (populate) работ, заказчиков и тестовой функции генерации имен
CTestTableFieldsReader.cls - тестовый класс для теста заполнения заказ-наряда
CTestWork.cls - тестовый класс для работ, используемых в заказ-наряде
CWorkComboBox.cls - реализация ComboBox для выбора работ
ICustomer.cls - интерфейс для Customer, пример имплементации - CTestCustomer
IDocumentTemplate.cls - интерфейс для генератора документов, пример имплементации CDefaultDocument
IFieldsReader.cls - интерфейс для ридера полей документа, примеры имплементации: CFieldsReaderFromXL, COrderFieldsReader
ILoaders.cls - интерфейс для загрузчиков значений Customers: loadCustomers, works: loadWorks, и функции генерации нумерации документов: getNewNumber
INameGenerator.cls - интерфейс для генератора имен документов, пример имплементации - CDefaultNameGenerator
ITables.cls - интерфейс для выбора таблиц, содержащих переменные поля
IWork.cls - интерфейс для работ. В данном случае работа включает наименование работы, наименование заменяемой запчасти, стоимость работы, цену запчасти
# СуммаПрописью.xlam
Надстройка для excel для указани суммы прописью в российских и белорусских рублях
## файлы
Tests.bas - тесты, примеры использования
СуммаПрописью.bas - реализация функций:
СуммаПрописью и СуммаРублейПрописью
# тестовые файлы
test_calling_form.xlsx - пример вызова формы из .xlsx файла
test_contract.docx - пример договора, заполняется из таблицы test_customer_card.xlsx
order_remplate.docx - пример заказ-наряда
test_doc_for_replacement - тестовый документ для тестирования поиска в файле и замены


















