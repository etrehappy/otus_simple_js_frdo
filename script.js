/////////////////////////////////////////////////
// Актуальный скрипт. Последнее обновление 04.02.2025
// Версии можно посмотреть здесь:
// https://github.com/etrehappy/otus_simple_js_frdo
/////////////////////////////////////////////////

// Register — это слово везде подразумевает основной лист, в котором заполняются данные. Например, лист 2025.
// ModalPage — это слово везде подразумевает лист, с которого забираются данные.
// GetRow() считает строки с 1, а не с 0

const RegSheetName = "2025";
const kSupportSheet = Api.GetSheet("Для 1ой линии");
const kSheetRegister = Api.GetSheet(RegSheetName);
const kSheetModalPage = Api.GetSheet("Заполнить УПК");
const kSheetUpkPdf = Api.GetSheet("PDF УПК");
const kSheetCheck = Api.GetSheet("Check");

const kCellForScriptStatus = kSupportSheet.GetRange("B1");
const kCellForScriptResult = kSupportSheet.GetRange("A10");
const kCellForScriptAttention = kSupportSheet.GetRange("A5");
const kCellForScriptAttentDescription = kSupportSheet.GetRange("A6");
const kDuplicateRowValueCell = kSupportSheet.GetRange("B3");
const kNameOfDuplicateRowCol = "BT";
const kMinRow = 3; // 4 в таблице
const kMaxRow = 2001; // 2002  в nf,kbwt
const kRegHeadersRowNumber = 2; // kSheetRegister
const kModalHeadersColNumber = 1; // kSheetModalPage
const kColumnForRowSearch = 1; // Поиск по столбцу "Вид документа" kSheetRegister
const kRegNumberFirstSymbol = '0'; 


const kPkType = "Повышение квалификации";
const kPpType = "Профессиональная переподготовка";
const kUpkKind = "Удостоверение о повышении квалификации";
const kDppKind = "Диплом о профессиональной переподготовке";
const kAllWorkCheckedTrue = "Да";
const kRusCitizenship = "РОССИЯ";
const kRusCitizenshipCode = "643";
const kFieldsNameIt = "Связь, информационные и коммуникационные технологии";
const kSpecGroupIt = "Информатика и вычислительная техника";


const kCurrentYearFirstDay = new Date(2025, 0, 1); // 1 января 2025
const kCurrentYearEndDay = new Date(2025, 11, 31); // 31 декабря 2025
const kEdLicenseDate = new Date(2018, 11, 28); //28 декабря 2018

const kUpkMinHours = 16;
const kDppMinHours = 250;

const kBgColorCustomError = Api.CreateColorFromRGB(255, 114, 107);
const kBgColorAttention = Api.CreateColorFromRGB(247, 202, 172);
const kBgColorScriptSuccess = Api.CreateColorFromRGB(169, 208, 141);
//const kBgColorCustomError = Api.CreateColorByName("deeppink");

const kRegexTabAndSpaces = /^[ \t\v\r\n\f]+|[ \t\v\r\n\f]+$/; // пробелы и табуляции в начале и в конце строки
const kRegexDppHoursNumberDot = /^\d+\.\s+/; // строки, начинающиеся с цифр, за которыми следуют точка и пробел (например, "1. " или "2. ").

const EDataSource = {
    kModal: 0
    , kReg: 1
};

// Ячейки страницы kSheetModalPage
const CellMp = {
    kDocKind: kSheetModalPage.GetRange("C2") // Ячейка "Вид документа"
    , kIssueDate: kSheetModalPage.GetRange("C3") // Ячейка "Дата выдачи документа"
    , kRegNumber: kSheetModalPage.GetRange("C4") // Ячейка "Рег. номер"
    , kCourseName: kSheetModalPage.GetRange("C5") // Ячейка "Наименование курса/программы"
    , kCourseStartDate: kSheetModalPage.GetRange("C6") // Ячейка "Дата начала обучения (на первом курсе)"
    , kCourseEndDate: kSheetModalPage.GetRange("C7") // Ячейка "Дата окончания обучения (на последнем курсе)"
    , kHoursSum: kSheetModalPage.GetRange("C8") // Ячейка "Срок обучения, часов (всего)"
    , kProject: kSheetModalPage.GetRange("C9") // Ячейка "Название итоговой работы (на последнем курсе)"

    , kLastName: kSheetModalPage.GetRange("C12") // Ячейка "Фамилия"
    , kFirstName: kSheetModalPage.GetRange("C13") // Ячейка "Имя"
    , kSecondName: kSheetModalPage.GetRange("C14") // Ячейка "Отчество"
    , kBirthdate: kSheetModalPage.GetRange("C15") // Ячейка "Дата рождения"
    , kSex: kSheetModalPage.GetRange("C16") // Ячейка "Пол"
    , kSnils: kSheetModalPage.GetRange("C17") // Ячейка "СНИЛС (обязательное только для РФ)"        
    , kCitizenship: kSheetModalPage.GetRange("C18") // Ячейка "Гражданство по ОКСМ (если нет СНИЛС)"                                                  
                                                  
    , kUniverDegree: kSheetModalPage.GetRange("C21") // Ячейка "Уровень образования по диплому о во или спо"
    , kSerialUniver: kSheetModalPage.GetRange("C22") // Ячейка "Серия диплома о во или спо"
    , kNumberUniver: kSheetModalPage.GetRange("C23") // Ячейка "Номер диплома о во или спо"
    , kUniverLastName: kSheetModalPage.GetRange("C24") // Ячейка "Фамилия указанная в дипломе о во или спо"                                                  
                                                  
    , kRefShtab: kSheetModalPage.GetRange("C27") // Ячейка "Ссылка на задачу в Shtab"
    , kSpecialist: kSheetModalPage.GetRange("C28") // Ячейка "Специалист 1л"
    , kComment: kSheetModalPage.GetRange("C29") // Ячейка "Комментарии"

    , kUniverIssueDate: kSheetModalPage.GetRange("C32") // Ячейка "Дата выдачи диплома о ВО/СПО"
    , kUniverRegNumber: kSheetModalPage.GetRange("C33") // Ячейка "Регистрационный номер диплома о ВО/СПО"
    , kDppQualifName: kSheetModalPage.GetRange("C34") // Ячейка "Название квалификации (ДПП)"
    , kSphereName: kSheetModalPage.GetRange("C35") // Ячейка "Название сферы из программы (ДПП)"
    , kPresentedCourses: kSheetModalPage.GetRange("C36") // Ячейка "Дисциплины, которые были защищены"
    , kHoursCourses: kSheetModalPage.GetRange("C37") // Ячейка "Количество часов по защищенным дисциплинам"
    , kAllWorksChecked: kSheetModalPage.GetRange("C38") // Ячейка "Внесены все дисциплины и защищены ПР на всех курсах (если есть)"
};

const CheckSheetStruct = {

    kColumnRegNumberList: 16 // Столбец "Регистрационные номера 2025" на листе Check
    , kCellRegNumber: null // если заполнен, то ячейка будет очищена в конце
    , kRegNumberMaxRow: 2000 // подготовленное кол-во рег.номеров в 2025 году
    , kCitizenshipMaxRow: 255 // в таблице 254
    , kColumnCitizenshipName: 9
    , kColumnCitizenshipId: 10         
};

// Столбцы реестра
const EColumnReg = {
  kDocKind: 1 // Столбец "Вид документа"
  
  /* //Не используется
  , kDocStatus: 2 // Столбец "Статус документа" (оригинал/дубликат)
  , kDocLoss: 3 // Столбец "Подтверждение утраты"
  , kDocChange: 4 // Столбец "Подтверждение обмена"
  , kDocDestruction: 5 // Столбец "Подтверждение уничтожения"
  , kDocSerial: 6 // Столбец "Серия документа"
  , kDocNumber: 7 // Столбец "Номер документа"
  */

  , kIssueDate: 8 // Столбец "Дата выдачи документа"
  , kRegNumber: 9 // Столбец "Рег. номер"
  , kProgramType: 10 // Столбец "Дополнительная профессиональная программа (повышение квалификации/профессиональная переподготовка)"
  , kCourseName: 11 // Столбец "Наименование доп. проф. программы"
  
   //На 02.2025 одинаковое для всех
  , kFieldsName: 12 // Столбец "Наименование области профессиональной деятельности"
  , kSpecGroup: 13 // Столбец "Укрупненные группы специальностей"
  , kQualificationName: 14 // Столбец "Наименование квалификации, профессии, специальности"
  , kUniverDegree: 15 // Столбец "Уровень образования"
  , kUniverLastName: 16 // Столбец "Фамилия указанная в дипломе о ВО или СПО"
  , kSerialUniver: 17 // Столбец "Серия диплома о ВО"
  , kNumberUniver: 18 // Столбец "Номер диплома о ВО"
  , kCourseStartYear: 19 // Столбец "Год начала"
  , kCourseEndYear: 20 // Столбец "Год окончания"
  , kHoursSum: 21 // Столбец "Срок обучения, часов (всего)"
  , kLastName: 22 // Столбец "Фамилия"
  , kFirstName: 23 // Столбец "Имя"
  , kSecondName: 24 // Столбец "Отчество"
  , kBirthdate: 25 // Столбец "Дата рождения"
  , kSex: 26 // Столбец "Пол"
  , kSnils: 27 // Столбец "СНИЛС"

  /* //Не используется
  , kTrainingForm: 28 // Столбец "Форма обучения"
  , kSourceFunding: 29 // Столбец "Источник финансирования обучения"
  , kEducationForm: 30 // Столбец "Форма получения образования"
  */

  , kCitizenshipId: 31 // Столбец "код страны по ОКСМ"

  /* //Не используется
  , kDocKindDupl: 32 // Столбец дубликата "Наименование документа (оригинала)"
  , kDocSerialDupl: 33 // Столбец дубликата "Серия (оригинала)"
  , kDocNumberDupl: 34 // Столбец дубликата "Номер (оригинала)"
  , kRegNumberDupl: 35 // Столбец дубликата "Регистрационный N (оригинала"
  , kIssueDateDupl: 36 // Столбец дубликата "Дата выдачи (оригинала)"
  , kLastNameDupl: 37 // Столбец дубликата "Фамилия получателя (оригинала)"
  , kFirstNameDupl: 38 // Столбец дубликата "Имя получателя (оригинала)"
  , kSecondNameDupl: 39 // Столбец дубликата "Отчество получателя (оригинала)"
  , kDocNumberForChange: 40 // Столбец "Номер документа для изменения"
  */
 
  , kRefShtab: 53 // Столбец "Ссылка на Shtab"
  , kSpecialist: 54 // Столбец "Специалист"
  , kComment: 55 // Столбец "Комментарии"
  , kCitizenship: 56 // Столбец "Гражданство"
  , kCourseStartDate: 57 // Столбец "Дата начала"
  , kCourseEndDate: 58 // Столбец "Дата окончания обучения (на последнем курсе)"
  , kProject: 59 // Столбец "Проектная работа"
  , kDppQualifName: 60 // Столбец "Название квалификации"
  , kSphereName: 61 // Столбец "Название сферы из программы" 
  , kUniverIssueDate: 62 // Столбец "Дата выдачи диплома о ВО"
  , kPresentedCourses: 63 // Столбец "Дисциплины, которые были защищены"
  , kHoursCourses: 64 // Столбец "Количество часов по защищенным дисциплинам"
  , kUniverRegNumber: 65 // Столбец "Регистрационный номер диплома о ВО"
  , kAllWorksChecked: 66 // Столбец "Внесены все дисциплины и защищены ПР на всех курсах (если есть)"
  , kEnterDataDate: 67 // Столбец "Дата внесения данных"
  , kChangeDataDate: 68 // Столбец "Дата выгрузки"
  , kAutoDuplicateSum: 70 // Столбец "Дата выдачи документа" 
  , kAutoDuplicateCheck: 71 // Столбец "Дата выдачи документа" 
  
};

const Text = {

  kScriptStart: "Скрипт запущен, жди ..."
  , kScriptCustomError: "Ошибка проверки"
  , kScriptWarning: "Предупреждение"
  , kScriptUnkonownError: "Неизвестная ошибка"
  , kScriptInfo: "Доп. информация"
  , kScriptAttention: "Внимание"
  , kScriptSuccess: "Данные успешно проверены или внесены в таблицу"

  , kErrorDateIssue: `Ошибка в дате выдачи. Она не может быть вне ${kCurrentYearFirstDay.getFullYear()} года или меньше даты окончания курса`
  , kErrorDocKind: "Вид документа, тип программы или поля ДПП не соответсвуют друг другу, либо не заполнены обязательные поля."
  , kErrorFailedDataFromCell: "Не удалось получить данные из ячейки"
  , kErrorFailedConfirmed: "Не получено подтверждение"
  , kErrorStartDate: "Нельзя выдавать УПК за курсы, которые начались до того, как была получена образовательная лицензия. Либо не указана дата начала обучения."
  , kErrorEndDate1: "Дата окончания курса не может быть меньше даты старта обучения."
  , kErrorEndDate2: `Дата окончания курса не может быть позднее ${kCurrentYearEndDay}.`
  , kErrorBirthdate: "Пользователю должно быть 18 лет или больше на момент окончания курса. Возраст пользователя:"
  , kErrorUniverIssueDate: "Дата выдачи диплома о ВО/СПО не может быть больше или равна дате выдачи ДПП."
  , kErrorDppHours: "Сумма часов отличается. Данные в столбцах должны совпадать."
  , kErrorAllWork: "Требуется подтвредить, что проектные работы на всех курсах защищены и указаны все дисциплины."
  , kErrorSnils: 
    `Некорректный ввод СНИЛС. Единый формат — ХХХ-ХХХ-ХХХ ХХ\n` + 
    `Если формат верный, проверь лишние пробелы в начале и конце.\n` + 
    `Текущее значение:`
  
   , kErrorFindRegNumber: "Не удалось найти регистрационный номер."
   , kErrorDuplicateRegNumber: "Первый символ рег. номер должен начинаться с 0." 
   , kErrorFindEmptyRow: "Не удалось найти свободную строку в реестре."
   , kErrorGetStartSheet: "Активный лист не учтён в скрипте."
   , kErrorCitizenshipId: "Не удалось определить Код страны по гражданству."
   , kErrorDuplicateRow: "В таблице дублируется строка. Проверь, что это не ошибка."
   , kErrorSnilsEmpty: "Заполнено гражданство «Россия», но нет СНИЛС."
   
   , kAttentionWithoutSecondName: "Не удалось получить Отчество, либо ячейка не заполнена."
   , kAttentionWithoutCitizenship: "Не задано поле Гражданство. Столбцы 31 и 56 заполнены автоматически. Убедись, что у слушатель есть российское гражданство. "
   , kAttentionWithoutSerialEd: "Отсутсвует серия диплома о ВО/СПО. Убедись, что её нет в дипломе."
   , kAttentionWithoutSex: "Проверь, что Пол указан корректно —"
   , kAttentionNotMatchSex: "Неудачная проверка пола. Проверь отчество. Убедись, что Пол указан корректно —"
   
   
};

const UpkPdfRange = {
    kRegNumber: kSheetUpkPdf.GetRange("AD27")
    , kIssueDate: kSheetUpkPdf.GetRange("AD35")
    , kLastName: kSheetUpkPdf.GetRange("CQ5")
    , kFirstAnsSecondName: kSheetUpkPdf.GetRange("CQ6")    
    , kStartDate: kSheetUpkPdf.GetRange("CK8")
    , kEndDate: kSheetUpkPdf.GetRange("CX8")
    , kCourseName: kSheetUpkPdf.GetRange("BV18")
    , kHoursSum: kSheetUpkPdf.GetRange("CQ22")
    , kProject: kSheetUpkPdf.GetRange("BT27")

};
const kUpkPdfCourseNameRange = kSheetUpkPdf.GetRange("BV18:DL19");
const kUpkPdfProjectRange = kSheetUpkPdf.GetRange("BT27:DO31");

Object.freeze(EColumnReg);
Object.freeze(Text);
Object.freeze(UpkPdfRange);
Object.freeze(CellMp);
Object.freeze(EDataSource);

/////////////////////////////////////////////////
// Классы
/////////////////////////////////////////////////

/* Не используется */
class Warning extends Error 
{
    constructor(message) 
    {
        super(message);
        this.name = "Warning";
    }
}

class CustomError extends Error 
{
    constructor(message) {
        super(message);
        this.name = "CustomError";
    }
}

class UserData 
{
    constructor(data_source, row = 0) 
    {   
        if(data_source === EDataSource.kModal)
        {
            this.#InitialisationWithModalData();
        }
        else if (data_source === EDataSource.kReg)
        {
            this.#InitialisationWithRegData(row);
        }
    }
    
    #InitialisationWithModalData() 
    {
        /* Курс */
        this.doc_kind_ = GetValueFromModal(CellMp.kDocKind);        
        this.issue_date_ = GetDateFromModal(CellMp.kIssueDate);
        this.reg_number_ = FindRegNumber(CellMp.kRegNumber);
        this.course_name_ = GetValueFromModal(CellMp.kCourseName);
        this.start_course_date_ = GetDateFromModal(CellMp.kCourseStartDate);
        this.end_course_date_ = GetDateFromModal(CellMp.kCourseEndDate);
        this.hours_sum_ = GetValueFromModal(CellMp.kHoursSum);
        this.project_ = GetValueFromModal(CellMp.kProject);

        /* Слушатель */        
        this.last_name_ = GetValueFromModal(CellMp.kLastName);
        this.first_name_ = GetValueFromModal(CellMp.kFirstName);
        this.second_name_ = CellMp.kSecondName.GetValue2(); //не нужно проверять на пустоту при инициализации 
        this.birthdate_ = GetDateFromModal(CellMp.kBirthdate);
        this.sex_ = GetValueFromModal(CellMp.kSex); 
        this.snils_ = CellMp.kSnils.GetValue2(); //не нужно проверять на пустоту при инициализации 
        this.citizenship_ = CellMp.kCitizenship.GetValue2(); //не нужно проверять на пустоту при инициализации         

        /* Диплом об образовании*/
        this.univer_degree_ = GetValueFromModal(CellMp.kUniverDegree); 
        this.serial_univer_ = CellMp.kSerialUniver.GetValue2(); //не нужно проверять на пустоту при инициализации 
        this.number_univer_ = GetValueFromModal(CellMp.kNumberUniver);
        this.univer_last_name_ = CellMp.kUniverLastName.GetValue2(); 
        
        /* Дополнительно*/
        this.ref_shtab_ = GetValueFromModal(CellMp.kRefShtab);
        this.specialist_ = GetValueFromModal(CellMp.kSpecialist);        
        this.comment_ = CellMp.kComment.GetValue(); 

        /* ДПП */
        this.univer_date_ = SerialDateToDate(CellMp.kUniverIssueDate.GetValue2()); //проверяется в CheckUniverIssueDate() 
        this.univer_reg_number_ = CellMp.kUniverRegNumber.GetValue2(); //не нужно проверять на пустоту при инициализации
        this.dpp_qualification_name_ = CellMp.kDppQualifName.GetValue2(); //проверяется на заполненность в CheckKindDocAndProgramType
        this.sphere_name_ = CellMp.kSphereName.GetValue2(); //не нужно проверять на пустоту при инициализации 
        this.presented_courses_ = CellMp.kPresentedCourses.GetValue(); //не нужно проверять на пустоту при инициализации 
        this.hours_courses_ = CellMp.kHoursCourses.GetValue(); // сравниваются с hours_sum_ в GetFormattedStringArray; не нужно проверять на пустоту при инициализации        
        this.all_works_checked_ = CellMp.kAllWorksChecked.GetValue2(); // не нужно проверять на пустоту при инициализации 

        /* Нет на странице */
        this.citizenship_id_ = null; // вычисляется в CheckDataFromModal => #FindCitizenshipId
        this.dpp_courses_and_hours_ = null; // вычисляется в CheckDataFromModal

    }

    #InitialisationWithRegData(row) 
    {
        this.doc_kind_ = GetValueFromReg(EColumnReg.kDocKind, row);
        this.issue_date_ = GetDateFromReg(EColumnReg.kIssueDate, row);
        this.reg_number_ = FindRegNumber (kSheetRegister.GetRangeByNumber(row, EColumnReg.kRegNumber) ); 
        this.course_name_ = GetValueFromReg(EColumnReg.kCourseName, row);
        this.univer_degree_ = GetValueFromReg(EColumnReg.kUniverDegree, row); 
        this.univer_last_name_ = kSheetRegister.GetRangeByNumber(row, EColumnReg.kUniverLastName).GetValue2(); //не нужно проверять на пустоту при инициализации 
        this.serial_univer_ = kSheetRegister.GetRangeByNumber(row, EColumnReg.kSerialUniver ).GetValue2(); //не нужно проверять на пустоту при инициализации 
        this.number_univer_ = GetValueFromReg(EColumnReg.kNumberUniver, row);
        this.hours_sum_ = GetValueFromReg(EColumnReg.kHoursSum, row);
        this.last_name_ = GetValueFromReg(EColumnReg.kLastName, row); 
        this.first_name_ = GetValueFromReg(EColumnReg.kFirstName, row);
        this.second_name_ = kSheetRegister.GetRangeByNumber(row, EColumnReg.kSecondName).GetValue2(); //не нужно проверять на пустоту при инициализации 
        this.birthdate_ = GetDateFromReg(EColumnReg.kBirthdate, row);
        this.sex_ = GetValueFromReg(EColumnReg.kSex, row);
        this.snils_ = kSheetRegister.GetRangeByNumber(row, EColumnReg.kSnils).GetValue2(); //не нужно проверять на пустоту при инициализации 
        this.ref_shtab_ = GetValueFromReg(EColumnReg.kRefShtab, row);
        this.specialist_ = GetValueFromReg(EColumnReg.kSpecialist, row);
        this.citizenship_ = kSheetRegister.GetRangeByNumber(row, EColumnReg.kCitizenship).GetValue2(); //не нужно проверять на пустоту при инициализации 
        this.start_course_date_ = GetDateFromReg(EColumnReg.kCourseStartDate, row);
        this.end_course_date_ = GetDateFromReg(EColumnReg.kCourseEndDate, row);
        this.project_ = GetValueFromReg(EColumnReg.kProject, row);
        this.dpp_qualification_name_ = kSheetRegister.GetRangeByNumber(row, EColumnReg.kDppQualifName).GetValue2(); //проверяется на заполненность в CheckKindDocAndProgramType
        this.sphere_name_ = kSheetRegister.GetRangeByNumber(row, EColumnReg.kSphereName).GetValue2(); //не нужно проверять на пустоту при инициализации 
        this.univer_date_ = SerialDateToDate(kSheetRegister.GetRangeByNumber(row, EColumnReg.kUniverIssueDate).GetValue2()); //проверяется в CheckUniverIssueDate() 
        this.presented_courses_ = kSheetRegister.GetRangeByNumber(row, EColumnReg.kPresentedCourses).GetValue(); //не нужно проверять на пустоту при инициализации 
        this.hours_courses_ = kSheetRegister.GetRangeByNumber(row, EColumnReg.kHoursCourses).GetValue(); // сравниваются с hours_sum_ в GetFormattedStringArray; не нужно проверять на пустоту при инициализации 
        this.univer_reg_number_ = kSheetRegister.GetRangeByNumber(row, EColumnReg.kUniverRegNumber).GetValue2(); //не нужно проверять на пустоту при инициализации 
        this.all_works_checked_ = kSheetRegister.GetRangeByNumber(row, EColumnReg.kAllWorksChecked).GetValue2(); // не нужно проверять на пустоту при инициализации 

        /* Нет на странице */
        this.citizenship_id_ = null; // вычисляется в CheckDataFromModal => #FindCitizenshipId
        this.dpp_courses_and_hours_ = null; // вычисляется в CheckDataFromModal

    }

      
}

class Verifier 
{
//public:
    constructor(data, data_source) 
    {     
        if (!(data instanceof UserData)) 
        {
            throw new Error("Аргумент должен быть объектом класса UserData");
        }

        this.data_ref_ = data;
        this.data_source_ = data_source;
    }

    #CheckRegNumberFirstSymbol()
    {
        if (!this.data_ref_.reg_number_.startsWith(kRegNumberFirstSymbol) ) 
            {
                throw new CustomError(Text.kErrorDuplicateRegNumber);
            }
    }

    #CheckRegNumberModal()
    {
        CheckSpacesModal(this.data_ref_.reg_number_, CellMp.kRegNumber);
        this.#CheckRegNumberFirstSymbol();
    }

    #CheckRegNumberFromReg(row)
    {  
        CheckSpacesReg(this.data_ref_.reg_number_, EColumnReg.kRegNumber, row);
        this.#CheckRegNumberFirstSymbol();
    }

    CheckDataFromModal() 
    {        
        this.#CheckDocKind(CellMp.kDocKind); //kSheetRegister.GetRangeByNumber(row, EColumnReg.kDocKind)
        this.#CheckRegNumberModal();
        this.#CheckIssueDate();
        this.#CheckTextDataModal();
        this.#CheckCourseDates();
        this.#CheckCourseHours();
        this.#CheckBirthdate();
        this.#CheckSex();
        /*
            univer_degree_ — проверятеся через GetValueFromModal при получении
            univer_last_name_ — проверятеся на заполненность при вставке данных в таблицу

            ref_shtab_ — проверятеся через GetValueFromModal при получении
            specialist_ — проверятеся через GetValueFromModal при получении
            comment_ — проверятеся на заполненность при вставке данных в таблицу
        */

        if (this.data_ref_.doc_kind_ === kDppKind)
        {
            CheckEmptyModalCell(this.data_ref_.univer_date_, CellMp.kUniverIssueDate); //т.к. не проверялось в конструкторе UserData

            this.#CheckUniverIssueDate();
            this.#CheckAllWorksCol(); 
            
            const adress_hours_dpp = CellMp.kHoursCourses.GetAddress(false, false, "xlA1");
            const adress_courses = CellMp.kPresentedCourses.GetAddress(false, false, "xlA1");
            const adress_hours_sum = CellMp.kHoursSum.GetAddress(false, false, "xlA1");  
            this.data_ref_.dpp_courses_and_hours_ = this.#GetDppCoursesAndHours(this.data_ref_, adress_hours_dpp, adress_courses, adress_hours_sum);    
        }

        if(this.data_ref_.snils_ )
        {
            this.#CheckSnils();

            if(this.data_ref_.citizenship_) // есть снилс и любое гражданство
            {
                this.data_ref_.citizenship_id_ = this.#FindCitizenshipId(this.data_ref_.citizenship_);
            }
            else // есть снилс, но не указано гражданство
            {             
                SetAttention();
                AddToAttentDescription(Text.kAttentionWithoutCitizenship); 

                this.data_ref_.citizenship_id_ = kRusCitizenshipCode; 
            }          
        }
        else if (this.data_ref_.citizenship_ === kRusCitizenship) // нет снилс, но есть российское гражданство
        {
            throw new CustomError(Text.kErrorSnilsEmpty);
        }
        else // нет снилс и российского гражданства
        {          
            CheckEmptyModalCell(this.data_ref_.citizenship_, CellMp.kCitizenship);
            this.data_ref_.citizenship_id_ = this.#FindCitizenshipId(this.data_ref_.citizenship_);
        }


       
    }

    CheckDataFromReg(row) 
    {
        const cell_doc_kind = kSheetRegister.GetRangeByNumber(row, EColumnReg.kDocKind);

        this.#CheckDocKind(cell_doc_kind);
        this.#CheckRegNumberFromReg(row);
        this.#CheckIssueDate();
        this.#CheckTextDataReg(row);
        this.#CheckCourseDates();
        this.#CheckCourseHours();
        this.#CheckBirthdate();
        this.#CheckSex();
        /*
            univer_degree_ — проверятеся через GetValueFromReg при получении
            univer_last_name_ — проверятеся на заполненность при вставке данных в таблицу 

            ref_shtab_ — проверятеся через GetValueFromReg при получении
            specialist_ — проверятеся через GetValueFromReg при получении
        */

        if (this.data_ref_.doc_kind_ === kDppKind)
        {
            CheckEmptyRegCell(this.data_ref_.univer_date_, EColumnReg.kUniverIssueDate, row); //т.к. не проверялось в конструкторе UserData
            this.#CheckUniverIssueDate();
            this.#CheckAllWorksCol(); 
            
            const adress_hours_dpp = kSheetRegister.GetRangeByNumber(row, EColumnReg.kHoursCourses).GetAddress(false, false, "xlA1");
            const adress_courses = kSheetRegister.GetRangeByNumber(row, EColumnReg.kPresentedCourses).GetAddress(false, false, "xlA1");
            const adress_hours_sum = kSheetRegister.GetRangeByNumber(row, EColumnReg.kHoursSum).GetAddress(false, false, "xlA1");  
            this.data_ref_.dpp_courses_and_hours_ = this.#GetDppCoursesAndHours(this.data_ref_, adress_hours_dpp, adress_courses, adress_hours_sum);    
        }

        if(this.data_ref_.snils_ )
        {
            this.#CheckSnils();

            if(this.data_ref_.citizenship_) // есть снилс и любое гражданство
            {
                this.data_ref_.citizenship_id_ = this.#FindCitizenshipId(this.data_ref_.citizenship_);
            }
            else // есть снилс, но не указано гражданство
            {             
                SetAttention();
                AddToAttentDescription(Text.kAttentionWithoutCitizenship); 

                this.data_ref_.citizenship_id_ = kRusCitizenshipCode; 
            }          
        }
        else if (this.data_ref_.citizenship_ === kRusCitizenship) // нет снилс, но есть российское гражданство
        {
            throw new CustomError(Text.kErrorSnilsEmpty);
        }
        else // нет снилс и российского гражданства
        {          
            CheckEmptyRegCell(this.data_ref_.citizenship_, EColumnReg.kCitizenship, row);
            this.data_ref_.citizenship_id_ = this.#FindCitizenshipId(this.data_ref_.citizenship_);
        }
       
    }
        
    
//private:

    #CheckDocKind(cell) 
    {
        
        if(this.data_ref_.doc_kind_ !== kUpkKind && this.data_ref_.doc_kind_ !== kDppKind)
        {
            const adress = cell.GetAddress(false, false, "xlA1");
            throw new CustomError(`Ошибка в ячейке ${adress}`);
        }
        
        if(this.data_ref_.doc_kind_ === kUpkKind && this.data_ref_.dpp_qualification_name_)
        {
            throw new CustomError(Text.kErrorDocKind);
        }
        else if ((this.data_ref_.doc_kind_ === kDppKind) && !this.data_ref_.dpp_qualification_name_)
        {
            /* В CheckTextDataModal (ниже) не проверяется qualification_name на заполненность. 
            CheckEmptyModalCell бросает исключение, поэтмоу не используется выше, 
            т.к. на этом шаге нужно просто проверить есть ли данные в dpp_qualification_name_ */

            throw new CustomError(Text.kErrorDocKind);
        }
    }

    #CheckIssueDate() 
    {
        const date_year = this.data_ref_.issue_date_.getFullYear();
        const current_year = kCurrentYearFirstDay.getFullYear();

        if( (this.data_ref_.issue_date_ < this.data_ref_.end_course_date_) || (date_year !== current_year))
        {
            throw new CustomError(Text.kErrorDateIssue);
        }
    }

    #CheckTextDataModal() 
    {
        CheckSpacesModal(this.data_ref_.number_univer_, CellMp.kNumberUniver);
        CheckSpacesModal(this.data_ref_.course_name_, CellMp.kCourseName);
        CheckSpacesModal(this.data_ref_.last_name_, CellMp.kLastName);
        CheckSpacesModal(this.data_ref_.first_name_, CellMp.kFirstName);
        CheckSpacesModal(this.data_ref_.project_, CellMp.kProject);

        if(this.data_ref_.second_name_)
        {
            CheckSpacesModal(this.data_ref_.second_name_, CellMp.kProject);
        }
        else
        {           
            SetAttention();
            AddToAttentDescription(Text.kErrorWithoutSecondName); 
        }

        if(this.data_ref_.doc_kind_ === kDppKind)
        {        
            CheckEmptyModalCell(this.data_ref_.sphere_name_, CellMp.kSphereName);
            CheckEmptyModalCell(this.data_ref_.presented_courses_, CellMp.kPresentedCourses);
            CheckEmptyModalCell(this.data_ref_.hours_courses_, CellMp.kHoursCourses);
            CheckEmptyModalCell(this.data_ref_.univer_reg_number_, CellMp.kUniverRegNumber);
            /*CheckEmptyModalCell(this.data_ref_.dpp_qualification_name_, CellMp.kDppQualifName); // поле qualification_name проверяется на заполненность в CheckDocKind*/
            
            /* Проверки на табуляцию и пробелы ниже не треубется, так как все эти данные вносятся в шаблон ДПП вручную */
            // CheckSpacesModal(this.data_ref_.dpp_qualification_name_, CellMp.kDppQualifName);
            // CheckSpacesModal(this.data_ref_.sphere_name_, CellMp.kSphereName);
            // CheckSpacesModal(this.data_ref_.presented_courses_, CellMp.kPresentedCourses);
            // CheckSpacesModal(this.data_ref_.hours_courses_, CellMp.kHoursCourses);
            // CheckSpacesModal(this.data_ref_.univer_reg_number_, CellMp.kUniverRegNumber);

        }

        if(this.data_ref_.serial_univer_)
        {
            CheckSpacesModal(this.data_ref_.serial_univer_, CellMp.kSerialUniver);
        }
        else
        {
            this.#AddInfoWithoutSerialEd();
        }
    }

    #CheckTextDataReg(row) 
    {

        CheckSpacesReg(this.data_ref_.number_univer_, EColumnReg.kNumberUniver, row);
        CheckSpacesReg(this.data_ref_.course_name_, EColumnReg.kCourseName, row);
        CheckSpacesReg(this.data_ref_.last_name_, EColumnReg.kLastName, row);
        CheckSpacesReg(this.data_ref_.first_name_, EColumnReg.kFirstName, row);
        CheckSpacesReg(this.data_ref_.project_, EColumnReg.kProject, row);

        if(this.data_ref_.second_name_)
        {    
            CheckSpacesReg(this.data_ref_.second_name_, EColumnReg.kSecondName, row);
        }
        else
        {           
            SetAttention();
            AddToAttentDescription(Text.kErrorWithoutSecondName); 
        }

        if(this.data_ref_.doc_kind_ === kDppKind)
        {        
            CheckEmptyRegCell(this.data_ref_.sphere_name_, EColumnReg.kSphereName, row);
            CheckEmptyRegCell(this.data_ref_.presented_courses_, EColumnReg.kPresentedCourses, row);
            CheckEmptyRegCell(this.data_ref_.hours_courses_, EColumnReg.kHoursCourses, row);
            CheckEmptyRegCell(this.data_ref_.univer_reg_number_, EColumnReg.kUniverRegNumber, row);
            /*CheckEmptyRegCell(this.data_ref_.dpp_qualification_name_, EColumnReg.kDppQualifName, row); // поле qualification_name проверяется на заполненность в CheckKindDocAndProgramType*/
         
            /* Проверки на табуляцию и пробелы ниже не треубется, так как все эти данные вносятся в шаблон ДПП вручную */
            //CheckSpacesReg(this.data_ref_.dpp_qualification_name_, EColumnReg.kDppQualifName, row);
            //CheckSpacesReg(this.data_ref_.sphere_name, EColumnReg.kSphereName, row);
            //CheckSpacesReg(this.data_ref_.presented_courses, EColumnReg.kPresentedCourses, row);
            //CheckSpacesReg(this.data_ref_.hours_courses, EColumnReg.kHoursCourses, row);
            //CheckSpacesReg(this.data_ref_.univer_reg_number, EColumnReg.kUniverRegNumber, row);
        }

        if(this.data_ref_.serial_univer_)
        {
            CheckSpacesReg(this.data_ref_.serial_univer_, EColumnReg.kSerialUniver, row);
        }
        else
        {
            this.#AddInfoWithoutSerialEd();
        }
    }


    #AddInfoWithoutSerialEd() 
    {
        SetInfo();
        AddToAttentDescription(Text.kAttentionWithoutSerialEd);
    }

    #CheckCourseDates() 
    {
        if(this.data_ref_.start_course_date_ < kEdLicenseDate)
        {
            throw new CustomError(Text.kErrorStartDate);
        }
        
        if(this.data_ref_.end_course_date_ < this.data_ref_.start_course_date_)
        {
            throw new CustomError(Text.kErrorEndDate1);
        }
        
        if(this.data_ref_.end_course_date_ > kCurrentYearEndDay)
        {
            throw new CustomError(Text.kErrorEndDate2);
        }
    }

    #CheckCourseHours() 
    {   
        if(this.data_ref_.doc_kind_ === kUpkKind)
        {
            if(this.data_ref_.hours_sum_ < kUpkMinHours || this.data_ref_.hours_sum_ >= kDppMinHours)
            {
                throw new CustomError(`Кол-во часов за всю программу не соответсвует виду документа. Для  «${kUpkKind}» должно быть от ${kUpkMinHours} до ${kDppMinHours} часов.`);
            }
        }
        else if(this.data_ref_.doc_kind_ === kDppKind)
        {
            if(this.data_ref_.hours_sum_ < kDppMinHours)
            {
                throw new CustomError(`Кол-во часов за всю программу не соответсвует виду документа. Для  «${kDppKind}» должно быть не менее ${kDppMinHours} часов.`);
            }
        }
    }

    #CheckBirthdate() 
    {
        let client_age = this.data_ref_.end_course_date_.getFullYear() - this.data_ref_.birthdate_.getFullYear();

        // Прошел ли день рождения
        if (this.data_ref_.end_course_date_.getMonth() < this.data_ref_.birthdate_.getMonth() 
            || (this.data_ref_.end_course_date_.getMonth() === this.data_ref_.birthdate_.getMonth() 
                && this.data_ref_.end_course_date_.getDate() < this.data_ref_.birthdate_.getDate()) ) 
        {
            client_age--;
        }
        
        if (client_age < 18) 
        {
            throw new CustomError(`${Text.kErrorBirthdate} ${client_age}`);
        }
    }

    #CheckSex() 
    {         
        const male = "Муж";
        const female = "Жен";       

        // Проверка окончания отчества
        if(!this.data_ref_.second_name_)
        {
            const error_text = `${Text.kAttentionWithoutSex} ${this.data_ref_.sex_}.`;

            //SetAttention(); (выводится через CheckTextData...)
            AddToAttentDescription(error_text);             
        }
        else if (IsMale(this.data_ref_.second_name_) && this.data_ref_.sex_ === male) {}
        else if (IsFemale(this.data_ref_.second_name_) && this.data_ref_.sex_ === female) {}
        else
        {
            const error_text = `${Text.kAttentionNotMatchSex} ${this.data_ref_.sex_}.`;

            SetAttention();
            AddToAttentDescription(error_text);            
        }

    }

    /*!
    * @brief Проверяется только для ДПП
    */
    #CheckUniverIssueDate() 
    {  
        if(this.data_ref_.issue_date_ <= this.data_ref_.univer_date_)
        {
            throw new CustomError(Text.kErrorUniverIssueDate);
        }
    }

    /*!
    * @brief Проверяется только для ДПП
    */
    #CheckAllWorksCol() 
    {         
        //Првоеряется на заполненность        
        if (this.data_ref_.all_works_checked_ !== kAllWorkCheckedTrue)
        {
            throw new CustomError(Text.kErrorAllWork);
        }
    }

    /*!
    * @brief Проверяется, если есть СНИЛС
    */
    #CheckSnils() 
    {
        const dash = "-";
        const space = " ";
        const snils_max_lenght = 14;
        const first_dash_pos = 3; 
        const second_dash_pos = 7;
        const space_pos = 11;

        if (this.data_ref_.snils_.length !== snils_max_lenght 
            || this.data_ref_.snils_.charAt(first_dash_pos) !== dash
            || this.data_ref_.snils_.charAt(second_dash_pos) !== dash 
            || this.data_ref_.snils_.charAt(space_pos) !== space ) 
        {
            throw new CustomError(`${Text.kErrorSnils} ${this.data_ref_.snils_}`);
        }

    }

    /*!
    * @brief Если иностранный гражданин, то нужно найти ID страны, иначе бросит исключение.
    */  
    #FindCitizenshipId(citizenship)
    {
        let id = null;
        const start_row = 1;   

        for (let row = start_row; row <= CheckSheetStruct.kCitizenshipMaxRow; row++) 
        {
            let value = kSheetCheck.GetRangeByNumber(row, CheckSheetStruct.kColumnCitizenshipName).GetValue2(); 
            
            if (value === citizenship) 
            { 
                id = kSheetCheck.GetRangeByNumber(row, CheckSheetStruct.kColumnCitizenshipId).GetValue2();            
                return id;
            }   
        }

        // if id = null
        throw new CustomError(Text.kErrorCitizenshipId);    
    }

    #GetDppCoursesAndHours(data, adress_hours_dpp, adress_courses, adress_hours_sum) 
    {  
        const separator_symb = "\n";    
        const presented_courses_string_array = data.presented_courses_.split(separator_symb);
        const hours_courses_string_array = data.hours_courses_.split(separator_symb);
        const hours_sum = Number(data.hours_sum_);
            
        this.#CheckDppCoursesAndHoursCells(adress_hours_dpp, adress_courses, presented_courses_string_array, hours_courses_string_array);
        
        return this.#GetFormattedStringArray(adress_hours_dpp, adress_hours_sum, hours_sum, presented_courses_string_array, hours_courses_string_array);
    }

    #GetFormattedStringArray(adress_hours_dpp, adress_hours_sum, hours_sum, presented_courses_string_array, hours_courses_string_array)
    {
            
        let hours_sum_dpp = 0;
        let formatted_string_array = "";

        for (let i = 0; i < presented_courses_string_array.length; i++) 
        {
            // Удаляем нумерацию из строки с часами
            const sanitized_hours_string = hours_courses_string_array[i].replace(kRegexDppHoursNumberDot, "");
            
            // Получаем сумму для сравнения с другим столбцом часов
            hours_sum_dpp += Number(sanitized_hours_string); 

            formatted_string_array += `${presented_courses_string_array[i]} — ${sanitized_hours_string}\n`;
        }
        
        //Сравниваем часы из разных стобцов
        if(hours_sum_dpp != hours_sum )
        {
            throw new CustomError(`${Text.kErrorDppHours} В ячейке ${adress_hours_dpp} по сумме отдельных курсов — ${hours_sum_dpp} ч., в ячейке ${adress_hours_sum} — ${hours_sum}.`);
        }
            
        return formatted_string_array;
    }

    #CheckDppCoursesAndHoursCells(adress_hours_dpp, adress_courses, presented_courses_string_array, hours_courses_string_array) 
    {
        

        if (presented_courses_string_array.length !== hours_courses_string_array.length) 
        {        
            throw new CustomError(`Количество строк в ячейках ${adress_courses} и ${adress_hours_dpp} должно совпадать. Для переноса строк внутри ячейки используй сочетание клавиш Alt (на macOS это Option) + Enter.`);
        }


        for (let i = 0; i < presented_courses_string_array.length; i++) 
        {
            if (!kRegexDppHoursNumberDot.test(presented_courses_string_array[i])) 
            {
                throw new CustomError(`Строка "${presented_courses_string_array[i]}" в ячейке ${adress_courses} не соответствует ожидаемому формату «Номер. Курс» (например, «1. Android Developer. Basic»). Проверь наличие порядкового номера, точки и пробела.`);
            }
        }

        for (let i = 0; i < hours_courses_string_array.length; i++) 
        {
            if (!kRegexDppHoursNumberDot.test(hours_courses_string_array[i])) 
            {
                throw new CustomError(`Строка "${hours_courses_string_array[i]}" в ячейке ${adress_hours_dpp} не соответствует ожидаемому формату «Номер. Часы» (например, «1. 172»). Проверь наличие порядкового номера, точки и пробела.`);
            }
        }
    }

}


/////////////////////////////////////////////////
// Начало
/////////////////////////////////////////////////

(function main()
{
    try    
    {
        /* Подготовка */
        CleanCells();
        const data_source = GetDataSource();

        

        let data = null;        
        
        if(data_source === EDataSource.kModal)
        {
            let reg_next_empty_row = null;
            data = new UserData(data_source);
            let verifier = new Verifier(data, data_source);
            reg_next_empty_row = FindNextEmptyRegRow();

            /* Проверка */
            verifier.CheckDataFromModal();   
            
            /* Вывод */
            EnterDataInTableFromModal(data, reg_next_empty_row);
            CleanSheetModalPage();
            UpdateDuplicateRowCell(reg_next_empty_row);
        }
        else if (data_source === EDataSource.kReg)
        {
            let reg_active_row = null;
            reg_active_row = kSheetRegister.GetActiveCell().GetRow() - 1;            
            data = new UserData(data_source, reg_active_row);         
            let verifier = new Verifier(data, data_source);
            verifier.CheckDataFromReg(reg_active_row);
            EnterDataInTableFromReg(data, reg_active_row);
            UpdateDuplicateRowCell(reg_active_row); 
        }

        ClearCellInRegNumberList();        
        CollectDataForShtab(data); /* Вывод данных для Shtab */
        ScriptSuccess();       
        CollectDataForPdf(data); /* Вывод данных для PDF */        
        kSupportSheet.SetActive();            
        
    } 
    catch (error) 
    {     
        kSupportSheet.SetActive();

        if (error instanceof CustomError) 
        {
            kCellForScriptStatus.SetValue(Text.kScriptCustomError);
            kCellForScriptStatus.SetFillColor(kBgColorCustomError)
            
            kCellForScriptResult.SetValue(error.message);
        } 
        else if (error instanceof Warning) 
        {
            kCellForScriptStatus.SetValue(Text.kScriptWarning);
            kCellForScriptResult.SetValue(error.message);
        } 
        else 
        {
            kCellForScriptStatus.SetValue(Text.kScriptUnkonownError);
            kCellForScriptResult.SetValue(error.message);
        } 
    }    
    
})();


/////////////////////////////////////////////////
// Подстановка данных Shtab
/////////////////////////////////////////////////

function CollectDataForShtab(data) 
{
    if (!(data instanceof UserData)) 
    {
        throw new Error("Аргумент должен быть объектом класса UserData");
    }

    let collected_data;

    if(data.doc_kind_ === kUpkKind)
    {
        collected_data = GetUpkTemplate(data);
    }
    else if(data.doc_kind_ === kDppKind)
    {
        collected_data = GetDppTemplate(data);
    }

    kCellForScriptResult.SetValue(collected_data);       
}

function GetUpkTemplate(data)
{
    let fio;
    const issue_date = FormatDateToDmy(data.issue_date_);
    const start_date = FormatDateToDmy(data.start_course_date_);
    const end_date = FormatDateToDmy(data.end_course_date_);


    if(data.second_name_)
    {
        fio = data.last_name_ + " " + data.first_name_ + " " + data.second_name_;
    }
    else
    {
        fio = data.last_name_ + " " + data.first_name_;
    }   

    let template = `• Регистрационный номер: ${data.reg_number_}` 
            + `\n` + `• Дата выдачи УПК: ${issue_date}`
            + `\n` + `• ФИО: ${fio}` 
            + `\n` + `• Дата обучения: с ${start_date} по ${end_date}` 
            + `\n` + `• Название курса: ${data.course_name_}` 
            + `\n` + `• Кол-во часов: ${data.hours_sum_}` 
            + `\n` + `• Итоговая работа: ${data.project_}` ;    

    return template;
}

function GetDppTemplate(data)
{
    let fio;
    const issue_date = FormatDateToDmy(data.issue_date_);
    const start_date = FormatDateToDmy(data.start_course_date_);
    const end_date = FormatDateToDmy(data.end_course_date_);
    const univer_date = FormatDateToDmy(data.univer_date_);

    if(data.second_name_)
    {
        fio = data.last_name_ + " " + data.first_name_ + " " + data.second_name_;
    }
    else
    {
        fio = data.last_name_ + " " + data.first_name_;
    }   

    let template = `• Регистрационный номер: ${data.reg_number_}` 
            + `\n` + `• Дата выдачи ДПП: ${issue_date}`            
            + `\n` + `• ФИО: ${fio}` 
            + `\n` + `• Название программы: ${data.course_name_}`
            + `\n` + `• Кол-во часов (всего): ${data.hours_sum_}`

            + `\n` + `• Название квалификации: ${data.dpp_qualification_name_}`
            + `\n` + `• Название сферы: ${data.sphere_name_}`
            + `\n` + `• Данные прошлого диплома: `
            + `\n` + ` - Уровень: ${data.univer_degree_}`
            + `\n` + ` - Серия: ${data.serial_univer_}`
            + `\n` + ` - Номер: ${data.number_univer_}`
            + `\n` + ` - Дата выдачи: ${univer_date}`
            + `\n` + ` - Рег. №: ${data.univer_reg_number_}` 

            + `\n` + `• Дата обучения: с ${start_date} по ${end_date}`
            + `\n` + `• Итоговая работа на последнем курсе: ${data.project_}` 
            + `\n` + `• Защищенные дисциплины (название — часы): `
            + `\n` + `${data.dpp_courses_and_hours_}`
            ;    

    return template;
}






/////////////////////////////////////////////////
// Подстановка данных PDF
/////////////////////////////////////////////////

function CollectDataForPdf(data)
{

    if(data.doc_kind_ === kUpkKind)
    {
        UpkPdfRange.kRegNumber.SetValue(data.reg_number_);
        UpkPdfRange.kIssueDate.SetValue(FormatDateToD_month_Y(data.issue_date_) );
        UpkPdfRange.kLastName.SetValue(data.last_name_.toUpperCase());
        UpkPdfRange.kFirstAnsSecondName.SetValue(GetFirstAnsSecondNameString(data.first_name_, data.second_name_) );
        UpkPdfRange.kStartDate.SetValue(FormatDateToDmy(data.start_course_date_) );
        UpkPdfRange.kEndDate.SetValue(FormatDateToDmy(data.end_course_date_) );
        UpkPdfRange.kCourseName.SetValue("«" + data.course_name_ + "»");
        UpkPdfRange.kHoursSum.SetValue(data.hours_sum_);
        UpkPdfRange.kProject.SetValue("«" + data.project_ + "»");
    }    
}

function GetFirstAnsSecondNameString(first_name, second_name)
{
    let string = first_name;

    if(second_name)
    {
        string += " " + second_name;
    }

    return string;
}




/////////////////////////////////////////////////
// Вспомогательное
/////////////////////////////////////////////////


function GetProgramType(doc_kind)
{
    if(doc_kind === kUpkKind)
    {
        return kPkType;
    }
    else if (doc_kind === kDppKind)
    {
        return kPpType;
    }
    else
    {
        //првоеряется в CheckDocKind
    }
}

/*!
* @brief Если фамилия из диплона не указана, подставится актуальная фамилия.
*/
function GetEdLastName(data)
{
    if(data.univer_last_name_ === "")
    {
        return data.last_name_;
    }
    else
    {
        return data.univer_last_name_;
    }
}

function AddFormuls(row)
{
    const table_row = row+1;
    kSheetRegister.GetRangeByNumber(row, EColumnReg.kAutoDuplicateSum).SetValue(`=$L${table_row}&$W${table_row}&$X${table_row}&$Y${table_row}`);
    kSheetRegister.GetRangeByNumber(row, EColumnReg.kAutoDuplicateCheck).SetValue(`=ЕСЛИ($BS${table_row}="";"";СЧЁТЕСЛИ($BS$${kMinRow+1}:$BS$${kMaxRow};$BS${table_row})>1)`);
}

/*!
* @brief Перенос из kSheetModalPage в kSheetRegister.
*/
function EnterDataInTableFromModal(data, empty_row)
{
    const program_type = GetProgramType(data.doc_kind_);
    const univer_last_name = GetEdLastName(data);

    //В порядке столбцов в kSheetRegister

    kSheetRegister.GetRangeByNumber(empty_row, EColumnReg.kDocKind).SetValue(data.doc_kind_);
    
    /* со 2 по 7 не заполняется*/
    
    kSheetRegister.GetRangeByNumber(empty_row, EColumnReg.kIssueDate).SetValue(data.issue_date_);
    kSheetRegister.GetRangeByNumber(empty_row, EColumnReg.kRegNumber).SetValue(data.reg_number_);
    kSheetRegister.GetRangeByNumber(empty_row, EColumnReg.kProgramType).SetValue(program_type);
    kSheetRegister.GetRangeByNumber(empty_row, EColumnReg.kCourseName).SetValue(data.course_name_);
    kSheetRegister.GetRangeByNumber(empty_row, EColumnReg.kFieldsName).SetValue(kFieldsNameIt);
    kSheetRegister.GetRangeByNumber(empty_row, EColumnReg.kSpecGroup).SetValue(kSpecGroupIt);
    kSheetRegister.GetRangeByNumber(empty_row, EColumnReg.kUniverDegree).SetValue(data.univer_degree_);
    kSheetRegister.GetRangeByNumber(empty_row, EColumnReg.kUniverLastName).SetValue(univer_last_name);

    if(data.serial_univer_)
    {
        kSheetRegister.GetRangeByNumber(empty_row, EColumnReg.kSerialUniver).SetValue(data.serial_univer_);
    }

    kSheetRegister.GetRangeByNumber(empty_row, EColumnReg.kNumberUniver).SetValue(data.number_univer_);
    kSheetRegister.GetRangeByNumber(empty_row, EColumnReg.kCourseStartYear).SetValue(data.start_course_date_.getFullYear() );
    kSheetRegister.GetRangeByNumber(empty_row, EColumnReg.kCourseEndYear).SetValue(data.end_course_date_.getFullYear() );    
    kSheetRegister.GetRangeByNumber(empty_row, EColumnReg.kHoursSum).SetValue(data.hours_sum_);    
    kSheetRegister.GetRangeByNumber(empty_row, EColumnReg.kLastName).SetValue(data.last_name_);
    kSheetRegister.GetRangeByNumber(empty_row, EColumnReg.kFirstName).SetValue(data.first_name_);

    if(data.second_name_)
    {
        kSheetRegister.GetRangeByNumber(empty_row, EColumnReg.kSecondName).SetValue(data.second_name_);
    }

    kSheetRegister.GetRangeByNumber(empty_row, EColumnReg.kBirthdate).SetValue(data.birthdate_);
    kSheetRegister.GetRangeByNumber(empty_row, EColumnReg.kSex).SetValue(data.sex_);

    if(data.snils_)
    {
        kSheetRegister.GetRangeByNumber(empty_row, EColumnReg.kSnils).SetValue(data.snils_);
    }

    /* с 28 по 30 не заполняется*/

    kSheetRegister.GetRangeByNumber(empty_row, EColumnReg.kCitizenshipId).SetValue(data.citizenship_id_);

    /* с 32 по 51 не заполняется*/

    kSheetRegister.GetRangeByNumber(empty_row, EColumnReg.kRefShtab).SetValue(data.ref_shtab_);
    kSheetRegister.GetRangeByNumber(empty_row, EColumnReg.kSpecialist).SetValue(data.specialist_);

    if(data.comment_)
    {
        kSheetRegister.GetRangeByNumber(empty_row, EColumnReg.kComment).SetValue(data.comment_);
    }

    if(data.citizenship_)
    {
        kSheetRegister.GetRangeByNumber(empty_row, EColumnReg.kCitizenship).SetValue(data.citizenship_);
    }
    else
    {
        kSheetRegister.GetRangeByNumber(empty_row, EColumnReg.kCitizenship).SetValue(kRusCitizenship);
    }
    
    
    kSheetRegister.GetRangeByNumber(empty_row, EColumnReg.kCourseStartDate).SetValue(data.start_course_date_);
    kSheetRegister.GetRangeByNumber(empty_row, EColumnReg.kCourseEndDate).SetValue(data.end_course_date_);
    kSheetRegister.GetRangeByNumber(empty_row, EColumnReg.kProject).SetValue(data.project_);
    
    if(data.doc_kind_ === kDppKind)
    {
        kSheetRegister.GetRangeByNumber(empty_row, EColumnReg.kQualificationName).SetValue(data.dpp_qualification_name_);

        kSheetRegister.GetRangeByNumber(empty_row, EColumnReg.kUniverIssueDate).SetValue(data.univer_date_);
        kSheetRegister.GetRangeByNumber(empty_row, EColumnReg.kUniverRegNumber).SetValue(data.univer_reg_number_);
        kSheetRegister.GetRangeByNumber(empty_row, EColumnReg.kDppQualifName).SetValue(data.dpp_qualification_name_);
        kSheetRegister.GetRangeByNumber(empty_row, EColumnReg.kSphereName).SetValue(data.sphere_name_);
        kSheetRegister.GetRangeByNumber(empty_row, EColumnReg.kPresentedCourses).SetValue(data.presented_courses_);
        kSheetRegister.GetRangeByNumber(empty_row, EColumnReg.kHoursCourses).SetValue(data.hours_courses_);
        kSheetRegister.GetRangeByNumber(empty_row, EColumnReg.kAllWorksChecked).SetValue(data.all_works_checked_);
    }

    let today = new Date();    
    kSheetRegister.GetRangeByNumber(empty_row, EColumnReg.kEnterDataDate).SetValue(FormatDateToDmyHm(today));
    
    AddFormuls(empty_row);
}

/*!
* @brief Автоматическое дозаполнение таблицы.
*/
function EnterDataInTableFromReg(data, current_row)
{
    debugger;
    const program_type = GetProgramType(data.doc_kind_);
    const univer_last_name = GetEdLastName(data);

    kSheetRegister.GetRangeByNumber(current_row, EColumnReg.kRegNumber).SetValue(data.reg_number_);
    kSheetRegister.GetRangeByNumber(current_row, EColumnReg.kProgramType).SetValue(program_type);
    kSheetRegister.GetRangeByNumber(current_row, EColumnReg.kFieldsName).SetValue(kFieldsNameIt);
    kSheetRegister.GetRangeByNumber(current_row, EColumnReg.kSpecGroup).SetValue(kSpecGroupIt);
    kSheetRegister.GetRangeByNumber(current_row, EColumnReg.kUniverLastName).SetValue(univer_last_name);
    kSheetRegister.GetRangeByNumber(current_row, EColumnReg.kCourseStartYear).SetValue(data.start_course_date_.getFullYear() );
    kSheetRegister.GetRangeByNumber(current_row, EColumnReg.kCourseEndYear).SetValue(data.end_course_date_.getFullYear() );    
    kSheetRegister.GetRangeByNumber(current_row, EColumnReg.kCitizenshipId).SetValue(data.citizenship_id_);
    
    if(data.citizenship_)
    {
        kSheetRegister.GetRangeByNumber(current_row, EColumnReg.kCitizenship).SetValue(data.citizenship_);
    }
    else
    {
        kSheetRegister.GetRangeByNumber(current_row, EColumnReg.kCitizenship).SetValue(kRusCitizenship);
    }

    if(data.doc_kind_ === kDppKind)
    {
        kSheetRegister.GetRangeByNumber(current_row, EColumnReg.kQualificationName).SetValue(data.dpp_qualification_name_);

    }
    
    const enter_data_cell = kSheetRegister.GetRangeByNumber(current_row, EColumnReg.kEnterDataDate);

    if(!enter_data_cell.GetValue2() )
    {
        let today = new Date();    
        enter_data_cell.SetValue(FormatDateToDmyHm(today));
    }
    else
    {
        let today = new Date();    
        kSheetRegister.GetRangeByNumber(current_row, EColumnReg.kChangeDataDate).SetValue(FormatDateToDmyHm(today));
    }

    AddFormuls(current_row);    

}


/*!
* @brief Преобразует серийное число в объект Date
*/
function SerialDateToDate(serial) 
{
    var utc_days = Math.floor(serial - 25569);
    var utc_value = utc_days * 86400;
    var date_info = new Date(utc_value * 1000);
    return new Date(date_info.getFullYear(), date_info.getMonth(), date_info.getDate());
}

/*!
* @brief Поиск пробелов и табуляций через регулярное выражение (данные из kSheetRegister)
*/
function CheckSpacesReg(cell_value, col, row) 
{       
    if(kRegexTabAndSpaces.test(cell_value))
    {
        const adress = kSheetRegister.GetRangeByNumber(row, col).GetAddress(false, false, "xlA1");
        const header = kSheetRegister.GetRangeByNumber(kRegHeadersRowNumber, col).GetValue2();
        
        throw new CustomError(`В ячейке ${adress} — ${header} — есть лишние символы пробелов или табуляции.`);
    }
}

/*!
* @brief Поиск пробелов и табуляций через регулярное выражение (данные из kSheetModalPage)
*/
function CheckSpacesModal(cell_value, cell) 
{   
    const row = cell.GetRow() - 1;

    if(kRegexTabAndSpaces.test(cell_value))
    {
        const adress = cell.GetAddress(false, false, "xlA1");
        const header = kSheetModalPage.GetRangeByNumber(row, kModalHeadersColNumber).GetValue2();
        
        throw new CustomError(`В ячейке ${adress} — ${header} — есть лишние символы пробелов или табуляции.`);
    }
}

function CleanSupportSheet() 
{
    kCellForScriptStatus.Clear();
    kCellForScriptStatus.SetFontName("Liberation Serif");
    kCellForScriptStatus.SetFontSize(12);
    kCellForScriptStatus.SetValue(Text.kScriptStart);
    
    kCellForScriptResult.Clear();
    kCellForScriptResult.SetFontName("Liberation Serif");
    kCellForScriptResult.SetFontSize(12);
    kCellForScriptResult.SetWrap(true);
    
    kCellForScriptAttention.Clear();
    kCellForScriptAttention.SetFontName("Liberation Serif");
    kCellForScriptAttention.SetFontSize(12);
    kCellForScriptAttention.SetWrap(true);
    
    kCellForScriptAttentDescription.Clear();
    kCellForScriptAttentDescription.SetFontName("Liberation Serif");
    kCellForScriptAttentDescription.SetFontSize(12);
    kCellForScriptAttentDescription.SetWrap(true);

    kDuplicateRowValueCell.Clear();
    kDuplicateRowValueCell.SetFontName("Liberation Serif");
    kDuplicateRowValueCell.SetFontSize(11);
}

function CleanSheetUpkPdf() 
{
    // УПК
    for (let key in UpkPdfRange) 
    {
        if (UpkPdfRange[key] && (typeof UpkPdfRange[key].Clear === "function") ) 
        {
            UpkPdfRange[key].Clear();
            UpkPdfRange[key].SetFontName("Liberation Serif");
            UpkPdfRange[key].SetFontSize(12);
            UpkPdfRange[key].SetNumberFormat("@"); //text
            UpkPdfRange[key].SetAlignHorizontal("center");
        }
    } 

    //kIssueDate.SetNumberFormat("dd.MM.yyyy");
    UpkPdfRange.kLastName.SetBold(true);
    UpkPdfRange.kFirstAnsSecondName.SetBold(true);
    UpkPdfRange.kProject.SetFontSize(11);        
    UpkPdfRange.kProject.SetAlignVertical("top");
    UpkPdfRange.kProject.SetWrap(true);
    UpkPdfRange.kCourseName.SetAlignVertical("center");        
    UpkPdfRange.kCourseName.SetWrap(true);

    kUpkPdfCourseNameRange.Merge(false);
    kUpkPdfProjectRange.Merge(false);
}

function CleanSheetModalPage() 
{
    for (let key in CellMp) 
    {
        if (CellMp[key] && (typeof CellMp[key].Clear === "function") ) 
        {
            CellMp[key].Clear();
            CellMp[key].SetFontName("Liberation Serif");
            CellMp[key].SetFontSize(11);
            CellMp[key].SetNumberFormat("@"); //text
            CellMp[key].SetAlignHorizontal("center");
            CellMp[key].SetAlignVertical("center");
        }
    }
    
    CellMp.kDocKind.SetAlignHorizontal("left");
    CellMp.kIssueDate.SetNumberFormat("dd/mm/yyyy");
    CellMp.kCourseStartDate.SetNumberFormat("dd/mm/yyyy");
    CellMp.kCourseEndDate.SetNumberFormat("dd/mm/yyyy");
    CellMp.kHoursSum.SetNumberFormat("0"); // 123
    CellMp.kProject.SetAlignHorizontal("left"); 
    CellMp.kBirthdate.SetNumberFormat("dd/mm/yyyy");
    CellMp.kComment.SetAlignHorizontal("left");
    CellMp.kUniverIssueDate.SetNumberFormat("dd/mm/yyyy");
    CellMp.kDppQualifName.SetAlignHorizontal("left");
    CellMp.kSphereName.SetAlignHorizontal("left");
    CellMp.kPresentedCourses.SetAlignHorizontal("left");
    CellMp.kHoursCourses.SetAlignHorizontal("left");
}

/*!
* @brief Очистка основных ячеек для вывода/ввода
*/
function CleanCells() 
{
    CleanSupportSheet();
    CleanSheetUpkPdf();
    //CleanSheetModalPage();
}

/*!
* @brief Информация в ячейке, что сотрудник вышел из проверки
*/
function Cancel() 
{
    kCellForScriptStatus.SetValue(Text.kErrorFailedConfirmed);
    /* kSheetRegister.SetActive(); // не работает корректно, если будет асинхронно или в callback */
}

/*!
* @brief Поиск пустых ячеек (данные из kSheetRegister)
*/
function CheckEmptyRegCell(value, col, row) 
{
    if(!value)
    {        
        const adress = kSheetRegister.GetRangeByNumber(row, col).GetAddress(false, false, "xlA1");
        const header = kSheetRegister.GetRangeByNumber(kRegHeadersRowNumber, col).GetValue2();
        
        throw new CustomError(`${Text.kErrorFailedDataFromCell} ${adress} — ${header}`);
    }

}

/*!
* @brief Поиск пустых ячеек (данные из kSheetModalPage)
*/
function CheckEmptyModalCell(value, cell) 
{
    const row = cell.GetRow() - 1;

    if(!value)
    {
        const adress = cell.GetAddress(false, false, "xlA1");
        const header = kSheetModalPage.GetRangeByNumber(row, kModalHeadersColNumber).GetValue2();
        
        throw new CustomError(`${Text.kErrorFailedDataFromCell} ${adress} — ${header}`);
    }

}

/*!
* @brief Получить значение из строки реестра в формате даты
*/
function GetDateFromReg(col, row) 
{
    const value = kSheetRegister.GetRangeByNumber(row, col).GetValue2();
    
    CheckEmptyRegCell(value, col, row);
    
    return SerialDateToDate(value);
}

/*!
* @brief Получить значение из строки реестра
*/
function GetValueFromReg(col, row) 
{ 
    const value = kSheetRegister.GetRangeByNumber(row, col).GetValue2();
    
    CheckEmptyRegCell(value, col, row);
    
    return value;
}

/*!
* @brief Получить значение из "модального окна" в формате даты 
*/
function GetDateFromModal(cell)
{
    const value = cell.GetValue2();
    
    CheckEmptyModalCell(value, cell);
    
    return SerialDateToDate(value);
}

/*!
* @brief Получить значение из "модального окна"
*/
function GetValueFromModal(cell)
{
    const value = cell.GetValue2();
    
    CheckEmptyModalCell(value, cell);
    
    return value;
}


/*!
* @brief Проверка отчества на пол (муж)
*/
function IsMale(second_name) 
{
    /* Предполагается, что second_name - валидное */
    
    const end_ch = "ч";
    const end_ogly = "оглы";
    const end_oglu = "оглу";
    const end_uly = "улы";
    const end_uuly = "уулу";
    
    if(second_name.endsWith(end_ch) 
        || second_name.endsWith(end_ogly)
        || second_name.endsWith(end_oglu)
        || second_name.endsWith(end_uly)
        || second_name.endsWith(end_uuly) )
    {
      return true;  
    }
    
    return false; 
}

/*!
* @brief Проверка отчества на пол (жен)
*/
function IsFemale(second_name) 
{
    /* Предполагается, что second_name - валидное */
    
    const end_a = "а";
    const end_kyzy = "кызы";
    const end_gyzy = "гызы";
    
    if(second_name.endsWith(end_a) 
        || second_name.endsWith(end_kyzy)
        || second_name.endsWith(end_gyzy) )
    {
      return true;  
    }
  
    return false; 
}

/*!
* @brief Меняет заголовок доп. ячейки
*/
function SetInfo() 
{
    const current_value = kCellForScriptAttention.GetValue2();

    if (!current_value) 
    {
        kCellForScriptAttention.SetValue(Text.kScriptInfo);
    }    
}

/*!
* @brief Меняет заголовок доп. ячейки
*/
function SetAttention() 
{
    const current_value = kCellForScriptAttention.GetValue2();

    if (!current_value) 
    {
        kCellForScriptAttention.SetValue(Text.kScriptAttention);        
    }

    kCellForScriptAttention.SetFillColor(kBgColorAttention);    
}

/*!
* @brief Меняет описание доп. ячейки
*/
function AddToAttentDescription(text) 
{
    const old_text = kCellForScriptAttentDescription.GetValue2();
    let new_text;

    if(old_text)
    {
        new_text = `${old_text}\n\n` + `${text}`;
    }
    else
    {
        new_text = text;
    }
    

    kCellForScriptAttentDescription.SetValue(new_text);
}

/*!
* @brief Преобразует дату в формат ДД-ММ-ГГГГ
*/
function FormatDateToDmy(date) 
{
    const day = ("0" + date.getDate()).slice(-2); // Добавляет ведущий ноль, если день состоит из одной цифры
    const month = ("0" + (date.getMonth() + 1)).slice(-2); // Добавляет ведущий ноль, если месяц состоит из одной цифры (месяцы в JS начинаются с 0)
    const year = date.getFullYear();

    return `${day}.${month}.${year}`;
}

/*!
* @brief Преобразует дату в формат ДД месяц ГГГГ
*/
function FormatDateToD_month_Y(date) 
{
    const day = ("0" + date.getDate()).slice(-2); // Добавляет ведущий ноль, если день состоит из одной цифры
    const months = ["января", "февраля", "марта", "апреля", "мая", "июня", "июля", "августа", "сентября", "октября", "ноября", "декабря"];
    const month = months[date.getMonth()]; // Получает название месяца
    const year = date.getFullYear();

    return `${day} ${month} ${year}`;
}

/*!
* @brief Преобразует дату в формат ДД.ММ.ГГ ЧЧ.ММ
*/
function FormatDateToDmyHm(date) 
{
    const day = ("0" + date.getDate()).slice(-2); // Добавляет ведущий ноль, если день состоит из одной цифры
    const month = ("0" + (date.getMonth() + 1)).slice(-2); // Добавляет ведущий ноль, если месяц состоит из одной цифры (месяцы в JS начинаются с 0)
    const year = date.getFullYear();
    
    const hours = ("0" + date.getHours()).slice(-2); // Добавляет ведущий ноль, если часы состоят из одной цифры
    const minutes = ("0" + date.getMinutes()).slice(-2); // Добавляет ведущий ноль, если минуты состоят из одной цифры

    return `${day}.${month}.${year} ${hours}:${minutes}`
}

function ScriptSuccess() 
{
    kCellForScriptStatus.SetValue(Text.kScriptSuccess);
    kCellForScriptStatus.SetFillColor(kBgColorScriptSuccess);
    Api.Save();
}

/*!
* @brief Ищет заполненную ячейку в соооветсвующем столбце kSheetCheck
*/
function FindNextRegNumberCell()
{
    const start_row = 1; // т.к. 0 - заголовок    

    for (let row = start_row; row < CheckSheetStruct.kRegNumberMaxRow; row++) 
    { 
        const cell = kSheetCheck.GetRangeByNumber(row, CheckSheetStruct.kColumnRegNumberList); 
        
        if (cell.GetValue2()) 
        {
            return cell;
        }
    }

    return null;
}

/*!
* @brief Получает рег. номер из kSheetCheck
*/
function GetNumberFromRegList()
{
    CheckSheetStruct.kCellRegNumber = FindNextRegNumberCell();

    if(CheckSheetStruct.kCellRegNumber == null)
    {
        throw new CustomError(Text.kErrorFindRegNumber);
    }

   return CheckSheetStruct.kCellRegNumber.GetValue2();
}

/*!
* @brief Если заполняем дубликат, то рег. номер будет получен из kSheetModalPage,  
* иначе рег. номер возьмем из kSheetCheck
*/
function FindRegNumber(data_cell)
{
    let reg_number = data_cell.GetValue2();      

    if(!reg_number)
    {
        reg_number = GetNumberFromRegList();
    }

    return reg_number;
}

/*!
* @brief Поиск строки по первому столбцу
*/
function FindNextEmptyRegRow()
{
    const start_row = 3; // 4 в таблице
    const last_row = 2001; // 2000 в таблице
    let cell = kSheetRegister.GetRange("B3"); // непустая ячейка
    
    let row = start_row;

    for (; row <= last_row; row++) 
    {
        cell = kSheetRegister.GetRangeByNumber(row, kColumnForRowSearch); 
        
        if (cell.GetValue2() === '') { break; }
    }

    if (cell.GetValue2() === kUpkKind || cell.GetValue2() === kDppKind) 
    {
        throw new CustomError(Text.kErrorFindEmptyRow);
    } 

    return row;
}

/*!
* @brief 
*/
function GetDataSource()
{   
    let active_sheet_name = Api.GetActiveSheet().GetName();

    if (active_sheet_name === kSheetModalPage.GetName() )
    {
        return EDataSource.kModal;
    }
    else if (active_sheet_name === kSheetRegister.GetName() )
    {
        return EDataSource.kReg;
    }
    else
    {
        throw new CustomError(Text.kErrorGetStartSheet);
    }
    
}

/*!
* @brief Дешёвая проверка на дубликаты
*/
function UpdateDuplicateRowCell(row)
{   
    const table_row = row + 1;

    kDuplicateRowValueCell.SetValue(`='${RegSheetName}'!${kNameOfDuplicateRowCol}${table_row}`);

    //Вариант ниже отрабатывает быстрее, чем формула на листе

    //const value =  kSheetRegister.GetRangeByNumber(row, EColumnReg.kAutoDuplicateCheck).GetValue2();
    // if(value === "ИСТИНА" ||  value === "TRUE")
    // {        
    //     const error_text = `${Text.kErrorDuplicateRow}`;
    //     SetAttention();
    //     AddToAttentDescription(error_text); 
    // }

}

function ClearCellInRegNumberList()
{
    if(CheckSheetStruct.kCellRegNumber !== null)
    {
        CheckSheetStruct.kCellRegNumber.Clear();
    }
}