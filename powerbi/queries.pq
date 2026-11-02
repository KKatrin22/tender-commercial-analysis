// =====================================================================
// Запросы Power Query отчёта kp_expertise.pbix. Данные учебные.
// Каждый запрос начинается со строки «// ===== Имя запроса =====».
// Чтобы воспроизвести: Home → Advanced Editor, вставить код запроса с таким же именем.
// Порядок создания: pFolder, tMap, tRef, KP_A, KP_B, KP_V1, KP_V2, KP_G, KP_Clean, tContr, tVOR.
// =====================================================================


// ===== pFolder (параметр, Manage parameters → New parameter, тип Text) =====
// Путь к папке data/source/ со слэшем в конце. Пример значения:
// "C:\KP_Project\data\source\" meta [IsParameterQuery = true, Type = "Text", IsParameterQueryRequired = true]


// ===== tMap: соответствие названий в КП и кодов ведомости =====
let
    Source = Excel.Workbook(File.Contents(pFolder & "name_mapping.xlsx"), null, true),
    Соответствия_Sheet = Source{[Item = "Соответствия", Kind = "Sheet"]}[Data],
    #"Promoted Headers" = Table.PromoteHeaders(Соответствия_Sheet, [PromoteAllScalars = true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers", {{"Наименование_в_КП", type text}, {"Код", type text}})
in
    #"Changed Type"


// ===== tRef: справочник цен без НДС на 31.03.2026 =====
let
    Source = Excel.Workbook(File.Contents(pFolder & "reference_prices.xlsx"), null, true),
    Справочник_Sheet = Source{[Item = "Справочник", Kind = "Sheet"]}[Data],
    #"Removed Top Rows" = Table.Skip(Справочник_Sheet, 3),
    #"Promoted Headers" = Table.PromoteHeaders(#"Removed Top Rows", [PromoteAllScalars = true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers", {
        {"Код", type text}, {"Наименование", type text}, {"Ед. изм.", type text},
        {"Вид затрат", type text}, {"Состав цены", type text},
        {"Цена без НДС на 31.03.2026, руб.", type number}, {"Мин., руб.", type number},
        {"Макс., руб.", type number}, {"Договоров в базе", Int64.Type}, {"Дата актуализации", type date}}),
    #"Renamed Columns" = Table.RenameColumns(#"Changed Type", {
        {"Ед. изм.", "Ед_спр"}, {"Состав цены", "Состав_спр"},
        {"Цена без НДС на 31.03.2026, руб.", "Цена_спр"}})
in
    #"Renamed Columns"


// ===== KP_A: КП Подрядчика А (цены без НДС, состав цены в отдельной колонке) =====
let
    Source = Excel.Workbook(File.Contents(pFolder & "kp_A.xlsx"), null, true),
    КП_Sheet = Source{[Item = "КП", Kind = "Sheet"]}[Data],
    #"Removed Top Rows" = Table.Skip(КП_Sheet, 4),
    #"Promoted Headers" = Table.PromoteHeaders(#"Removed Top Rows", [PromoteAllScalars = true]),
    #"Filtered Rows" = Table.SelectRows(#"Promoted Headers", each [#"№"] <> null),
    #"Renamed Columns" = Table.RenameColumns(#"Filtered Rows", {
        {"№", "Поз"}, {"Ед. изм.", "Ед"}, {"Кол-во", "Кол_во"},
        {"Цена за ед., руб. без НДС", "Цена_в_КП"},
        {"Стоимость, руб. без НДС", "Сумма_в_КП"},
        {"Состав цены", "Состав_цены"}}),
    #"Changed Type" = Table.TransformColumnTypes(#"Renamed Columns", {
        {"Поз", Int64.Type}, {"Наименование", type text}, {"Ед", type text},
        {"Кол_во", type number}, {"Цена_в_КП", type number},
        {"Сумма_в_КП", type number}, {"Состав_цены", type text}}),
    #"Added Подрядчик" = Table.AddColumn(#"Changed Type", "Подрядчик", each "Подрядчик А", type text),
    #"Added ДатаКП" = Table.AddColumn(#"Added Подрядчик", "ДатаКП", each #date(2026, 2, 12), type date),
    #"Added Цена_без_НДС" = Table.AddColumn(#"Added ДатаКП", "Цена_без_НДС_руб", each [Цена_в_КП], type number),
    #"Added Сумма_без_НДС" = Table.AddColumn(#"Added Цена_без_НДС", "Сумма_без_НДС_руб", each [Сумма_в_КП], type number),
    #"Added Доставка" = Table.AddColumn(#"Added Сумма_без_НДС", "ДоставкаВключена", each "Нет", type text)
in
    #"Added Доставка"


// ===== KP_B: КП Подрядчика Б (цены с НДС 22%, состав цены в названии) =====
let
    Source = Excel.Workbook(File.Contents(pFolder & "kp_B.xlsx"), null, true),
    Лист1_Sheet = Source{[Item = "Лист1", Kind = "Sheet"]}[Data],
    #"Removed Top Rows" = Table.Skip(Лист1_Sheet, 4),
    #"Promoted Headers" = Table.PromoteHeaders(#"Removed Top Rows", [PromoteAllScalars = true]),
    #"Filtered Rows" = Table.SelectRows(#"Promoted Headers", each [#"Поз."] <> null),
    #"Renamed Columns" = Table.RenameColumns(#"Filtered Rows", {
        {"Поз.", "Поз"}, {"Ед.", "Ед"}, {"Кол-во", "Кол_во"},
        {"Цена за ед. с НДС, руб.", "Цена_в_КП"}, {"Сумма с НДС, руб.", "Сумма_в_КП"}}),
    #"Changed Type" = Table.TransformColumnTypes(#"Renamed Columns", {
        {"Поз", Int64.Type}, {"Наименование", type text}, {"Ед", type text},
        {"Кол_во", type number}, {"Цена_в_КП", type number}, {"Сумма_в_КП", type number}}),
    #"Added Подрядчик" = Table.AddColumn(#"Changed Type", "Подрядчик", each "Подрядчик Б", type text),
    #"Added ДатаКП" = Table.AddColumn(#"Added Подрядчик", "ДатаКП", each #date(2026, 2, 20), type date),
    #"Added Цена_без_НДС" = Table.AddColumn(#"Added ДатаКП", "Цена_без_НДС_руб", each [Цена_в_КП] / 1.22, type number),
    #"Added Сумма_без_НДС" = Table.AddColumn(#"Added Цена_без_НДС", "Сумма_без_НДС_руб", each [Сумма_в_КП] / 1.22, type number),
    #"Added Доставка" = Table.AddColumn(#"Added Сумма_без_НДС", "ДоставкаВключена", each "Нет", type text),
    #"Added Состав" = Table.AddColumn(#"Added Доставка", "Состав_цены", each
        if Text.Contains([Наименование], "с материал") then "работа с материалом"
        else if Text.Contains([Наименование], "бетон заказчика") then "работа"
        else "не указан", type text)
in
    #"Added Состав"


// ===== KP_V1: КП Подрядчика В, лист «Оборудование» (цены в тыс. руб. с НДС 20%, состав не указан) =====
let
    Source = Excel.Workbook(File.Contents(pFolder & "kp_V.xlsx"), null, true),
    Оборудование_Sheet = Source{[Item = "Оборудование", Kind = "Sheet"]}[Data],
    #"Removed Top Rows" = Table.Skip(Оборудование_Sheet, 4),
    #"Promoted Headers" = Table.PromoteHeaders(#"Removed Top Rows", [PromoteAllScalars = true]),
    #"Filtered Rows" = Table.SelectRows(#"Promoted Headers", each [#"№"] <> null),
    #"Renamed Columns" = Table.RenameColumns(#"Filtered Rows", {
        {"№", "Поз"}, {"Ед.", "Ед"}, {"Объём", "Кол_во"},
        {"Цена, тыс. руб. за ед. (с НДС 20%)", "Цена_в_КП"},
        {"Итого, тыс. руб. (с НДС 20%)", "Сумма_в_КП"}}),
    #"Changed Type" = Table.TransformColumnTypes(#"Renamed Columns", {
        {"Поз", Int64.Type}, {"Наименование", type text}, {"Ед", type text},
        {"Кол_во", type number}, {"Цена_в_КП", type number}, {"Сумма_в_КП", type number}}),
    #"Added Подрядчик" = Table.AddColumn(#"Changed Type", "Подрядчик", each "Подрядчик В", type text),
    #"Added ДатаКП" = Table.AddColumn(#"Added Подрядчик", "ДатаКП", each #date(2025, 11, 14), type date),
    #"Added Цена_без_НДС" = Table.AddColumn(#"Added ДатаКП", "Цена_без_НДС_руб", each [Цена_в_КП] * 1000 / 1.2, type number),
    #"Added Сумма_без_НДС" = Table.AddColumn(#"Added Цена_без_НДС", "Сумма_без_НДС_руб", each [Сумма_в_КП] * 1000 / 1.2, type number),
    #"Added Доставка" = Table.AddColumn(#"Added Сумма_без_НДС", "ДоставкаВключена", each "Нет", type text),
    #"Added Состав" = Table.AddColumn(#"Added Доставка", "Состав_цены", each "не указан", type text)
in
    #"Added Состав"


// ===== KP_V2: КП Подрядчика В, лист «Работы» (те же условия) =====
let
    Source = Excel.Workbook(File.Contents(pFolder & "kp_V.xlsx"), null, true),
    Работы_Sheet = Source{[Item = "Работы", Kind = "Sheet"]}[Data],
    #"Removed Top Rows" = Table.Skip(Работы_Sheet, 4),
    #"Promoted Headers" = Table.PromoteHeaders(#"Removed Top Rows", [PromoteAllScalars = true]),
    #"Filtered Rows" = Table.SelectRows(#"Promoted Headers", each [#"№"] <> null),
    #"Renamed Columns" = Table.RenameColumns(#"Filtered Rows", {
        {"№", "Поз"}, {"Ед.", "Ед"}, {"Объём", "Кол_во"},
        {"Цена, тыс. руб. за ед. (с НДС 20%)", "Цена_в_КП"},
        {"Итого, тыс. руб. (с НДС 20%)", "Сумма_в_КП"}}),
    #"Changed Type" = Table.TransformColumnTypes(#"Renamed Columns", {
        {"Поз", Int64.Type}, {"Наименование", type text}, {"Ед", type text},
        {"Кол_во", type number}, {"Цена_в_КП", type number}, {"Сумма_в_КП", type number}}),
    #"Added Подрядчик" = Table.AddColumn(#"Changed Type", "Подрядчик", each "Подрядчик В", type text),
    #"Added ДатаКП" = Table.AddColumn(#"Added Подрядчик", "ДатаКП", each #date(2025, 11, 14), type date),
    #"Added Цена_без_НДС" = Table.AddColumn(#"Added ДатаКП", "Цена_без_НДС_руб", each [Цена_в_КП] * 1000 / 1.2, type number),
    #"Added Сумма_без_НДС" = Table.AddColumn(#"Added Цена_без_НДС", "Сумма_без_НДС_руб", each [Сумма_в_КП] * 1000 / 1.2, type number),
    #"Added Доставка" = Table.AddColumn(#"Added Сумма_без_НДС", "ДоставкаВключена", each "Нет", type text),
    #"Added Состав" = Table.AddColumn(#"Added Доставка", "Состав_цены", each "не указан", type text)
in
    #"Added Состав"


// ===== KP_G: КП Подрядчика Г (поставка и ПНР, цены с доставкой и НДС 22%, «КРУ под ключ») =====
let
    Source = Excel.Workbook(File.Contents(pFolder & "kp_G.xlsx"), null, true),
    КП_Sheet = Source{[Item = "КП", Kind = "Sheet"]}[Data],
    #"Removed Top Rows" = Table.Skip(КП_Sheet, 4),
    #"Promoted Headers" = Table.PromoteHeaders(#"Removed Top Rows", [PromoteAllScalars = true]),
    #"Filtered Rows" = Table.SelectRows(#"Promoted Headers", each [#"№"] <> null),
    #"Renamed Columns" = Table.RenameColumns(#"Filtered Rows", {
        {"№", "Поз"}, {"Ед.", "Ед"}, {"Кол-во", "Кол_во"},
        {"Цена за ед. с доставкой и НДС 22%, руб.", "Цена_в_КП"},
        {"Сумма с НДС, руб.", "Сумма_в_КП"}, {"Что входит в цену", "Состав_в_КП"}}),
    #"Changed Type" = Table.TransformColumnTypes(#"Renamed Columns", {
        {"Поз", Int64.Type}, {"Наименование", type text}, {"Ед", type text},
        {"Кол_во", type number}, {"Цена_в_КП", type number},
        {"Сумма_в_КП", type number}, {"Состав_в_КП", type text}}),
    #"Added Подрядчик" = Table.AddColumn(#"Changed Type", "Подрядчик", each "Подрядчик Г", type text),
    #"Added ДатаКП" = Table.AddColumn(#"Added Подрядчик", "ДатаКП", each #date(2026, 2, 25), type date),
    #"Added Цена_без_НДС" = Table.AddColumn(#"Added ДатаКП", "Цена_без_НДС_руб", each [Цена_в_КП] / 1.22, type number),
    #"Added Сумма_без_НДС" = Table.AddColumn(#"Added Цена_без_НДС", "Сумма_без_НДС_руб", each [Сумма_в_КП] / 1.22, type number),
    #"Added Доставка" = Table.AddColumn(#"Added Сумма_без_НДС", "ДоставкаВключена", each "Да", type text),
    #"Added Состав" = Table.AddColumn(#"Added Доставка", "Состав_цены", each
        if Text.Contains([Состав_в_КП], "под ключ") then "под ключ"
        else if Text.StartsWith([Состав_в_КП], "оборудование") then "поставка"
        else [Состав_в_КП], type text),
    #"Removed Columns" = Table.RemoveColumns(#"Added Состав", {"Состав_в_КП"})
in
    #"Removed Columns"


// ===== KP_Clean: все КП в одной таблице, коды, сравнение со справочником, флаги =====
let
    Source = Table.Combine({KP_A, KP_B, KP_V1, KP_V2, KP_G}),
    #"Changed Type" = Table.TransformColumnTypes(Source, {
        {"Кол_во", type number}, {"Цена_в_КП", type number}, {"Сумма_в_КП", type number},
        {"Цена_без_НДС_руб", type number}, {"Сумма_без_НДС_руб", type number}}),
    // Ключ для сопоставления: название без пометок о составе цены
    #"Added Наим_ключ" = Table.AddColumn(#"Changed Type", "Наим_ключ", each
        Text.Trim(Text.Replace(Text.Replace([Наименование], " (с материалами)", ""), " (работа, бетон заказчика)", ""))),
    #"Merged tMap" = Table.NestedJoin(#"Added Наим_ключ", {"Наим_ключ"}, tMap, {"Наименование_в_КП"}, "tMap", JoinKind.LeftOuter),
    #"Expanded tMap" = Table.ExpandTableColumn(#"Merged tMap", "tMap", {"Код"}, {"Код"}),
    #"Merged tRef" = Table.NestedJoin(#"Expanded tMap", {"Код"}, tRef, {"Код"}, "tRef", JoinKind.LeftOuter),
    #"Expanded tRef" = Table.ExpandTableColumn(#"Merged tRef", "tRef", {"Ед_спр", "Состав_спр", "Цена_спр"}, {"Ед_спр", "Состав_спр", "Цена_спр"}),
    // Индексация к 31.03.2026, условный рост 8,5% в год
    #"Added Цена_инд_руб" = Table.AddColumn(#"Expanded tRef", "Цена_инд_руб", each
        [Цена_без_НДС_руб] * Number.Power(1.085, Duration.Days(#date(2026, 3, 31) - [ДатаКП]) / 365.25), type number),
    #"Added Откл_проц" = Table.AddColumn(#"Added Цена_инд_руб", "Откл_проц", each
        if [Цена_спр] = null then null else ([Цена_инд_руб] - [Цена_спр]) / [Цена_спр], type number),
    #"Added Флаг" = Table.AddColumn(#"Added Откл_проц", "Флаг", each
        if [Цена_спр] = null then "нет в справочнике"
        else if [Ед] <> [Ед_спр] then "ед. изм. не совпадает"
        else if [Состав_цены] <> "не указан" and [Состав_цены] <> [Состав_спр] then "другой состав цены"
        else if Number.Abs([Откл_проц]) > 0.15 then
            (if [Состав_цены] = "не указан" and [Состав_спр] <> "поставка" then "уточнить состав цены" else "отклонение больше 15%")
        else "в норме", type text),
    #"Added Сумма_инд_руб" = Table.AddColumn(#"Added Флаг", "Сумма_инд_руб", each
        [Сумма_без_НДС_руб] * Number.Power(1.085, Duration.Days(#date(2026, 3, 31) - [ДатаКП]) / 365.25), type number),
    #"Added Причина" = Table.AddColumn(#"Added Сумма_инд_руб", "Причина", each
        if [Флаг] = "отклонение больше 15%" then (if [Откл_проц] > 0 then "Цена выше справочной" else "Цена ниже справочной")
        else if [Флаг] = "ед. изм. не совпадает" then "Другая единица измерения"
        else if [Флаг] = "другой состав цены" then "Другой состав цены"
        else if [Флаг] = "уточнить состав цены" then "Состав цены не указан"
        else if [Флаг] = "нет в справочнике" then "Нет в справочнике"
        else "", type text),
    #"Added Что_сделать" = Table.AddColumn(#"Added Причина", "Что_сделать", each
        if [Флаг] = "отклонение больше 15%" and [Откл_проц] > 0 then "Запросить у подрядчика обоснование цены"
        else if [Флаг] = "отклонение больше 15%" then "Проверить тип и производителя: цена ниже справочной"
        else if [Флаг] = "ед. изм. не совпадает" then "Пересчитать цену в единицу ведомости"
        else if [Флаг] = "другой состав цены" then "Уточнить состав цены или добавить стоимость материала"
        else if [Флаг] = "уточнить состав цены" then "Запросить состав цены: включены ли материалы"
        else if [Флаг] = "нет в справочнике" then "Нет прямого аналога: сравнивать по сумме позиций ведомости"
        else "", type text),
    // У позиций с другой единицей процент отклонения бессмыслен, в отчёте не показываем
    #"Added Откл_для_показа" = Table.AddColumn(#"Added Что_сделать", "Откл_для_показа", each
        if [Флаг] = "ед. изм. не совпадает" then null else [Откл_проц], type number),
    #"Changed Type1" = Table.TransformColumnTypes(#"Added Откл_для_показа", {{"ДатаКП", type date}})
in
    #"Changed Type1"


// ===== tContr: список подрядчиков =====
Table.Distinct(Table.SelectColumns(KP_Clean, {"Подрядчик"}))


// ===== tVOR: ведомость × подрядчики, статус позиции, итоговая сумма =====
let
    Source = Excel.Workbook(File.Contents(pFolder & "vor_object.xlsx"), null, true),
    #"Ведомость объёмов_Sheet" = Source{[Item = "Ведомость объёмов", Kind = "Sheet"]}[Data],
    #"Removed Top Rows" = Table.Skip(#"Ведомость объёмов_Sheet", 4),
    #"Promoted Headers" = Table.PromoteHeaders(#"Removed Top Rows", [PromoteAllScalars = true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers", {
        {"№ п/п", Int64.Type}, {"Код позиции", type text}, {"Объект", type text},
        {"Вид затрат", type text}, {"Наименование", type text}, {"Ед. изм.", type text},
        {"Объём", type number}, {"Состав цены", type text}}),
    #"Renamed Columns" = Table.RenameColumns(#"Changed Type", {{"Код позиции", "Код"}, {"Вид затрат", "Вид_затрат"}}),
    // Каждая позиция ведомости повторяется для каждого подрядчика (45 × 4 = 180 строк)
    #"Added Подрядчик" = Table.AddColumn(#"Renamed Columns", "Подрядчик", each tContr[Подрядчик]),
    #"Expanded Подрядчик" = Table.ExpandListColumn(#"Added Подрядчик", "Подрядчик"),
    #"Merged KP_Clean" = Table.NestedJoin(#"Expanded Подрядчик", {"Код", "Подрядчик"}, KP_Clean, {"Код", "Подрядчик"}, "KP_Clean", JoinKind.LeftOuter),
    #"Expanded KP_Clean" = Table.ExpandTableColumn(#"Merged KP_Clean", "KP_Clean", {"Сумма_без_НДС_руб", "Сумма_инд_руб"}, {"Сумма_без_НДС_руб", "Сумма_инд_руб"}),
    // Статус: цена есть в КП / позиция входит в «КРУ под ключ» Подрядчика Г / цены нет
    #"Added Статус" = Table.AddColumn(#"Expanded KP_Clean", "Статус", each
        if [Сумма_без_НДС_руб] <> null then "есть"
        else if List.Contains({"P07", "P26", "P41"}, [Код]) and [Подрядчик] = "Подрядчик Г" then "в позиции под ключ"
        else "нет в КП", type text),
    #"Merged tRef" = Table.NestedJoin(#"Added Статус", {"Код"}, tRef, {"Код"}, "tRef", JoinKind.LeftOuter),
    #"Expanded tRef" = Table.ExpandTableColumn(#"Merged tRef", "tRef", {"Цена_спр"}, {"Цена_спр"}),
    #"Changed Type1" = Table.TransformColumnTypes(#"Expanded tRef", {{"Цена_спр", type number}}),
    // Итоговая сумма: цена КП (индексированная) / вся «КРУ под ключ» записана на P07 / справочная цена × объём
    #"Added Сумма_итог_руб" = Table.AddColumn(#"Changed Type1", "Сумма_итог_руб", each
        if [Статус] = "есть" then [Сумма_инд_руб]
        else if [Статус] = "в позиции под ключ" then
            (if [Код] = "P07" then List.Sum(Table.SelectRows(KP_Clean, each [Код] = "P07+P26+P41")[Сумма_инд_руб]) else 0)
        else [Объём] * [Цена_спр], type number),
    #"Rounded Off" = Table.TransformColumns(#"Added Сумма_итог_руб", {{"Сумма_итог_руб", each Number.Round(_, 2), type number}}),
    #"Added Источник суммы" = Table.AddColumn(#"Rounded Off", "Источник суммы", each
        if [Статус] = "нет в КП" then "По справочным ценам (нет в КП)" else "По ценам КП", type text)
in
    #"Added Источник суммы"
