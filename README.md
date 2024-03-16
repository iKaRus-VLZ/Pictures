# Pictures
Some examples of working with Picture data in VBA (Access)
База содержит примеры работы с изображениями в Фссуыы 
Ззагрузка/выгрузка файла изображения в таблицу для последующего использования,
вывод данных из таблицы на контролы типа Access.CommandButton, Access.Image и все принимающие StdPicture
+ некоторые дополнительные возможности последующей работы с загруженным изображением (добавление текста, повороты, полупрозрачность и т.п.)
поддерживает альфаканал, PNG и ICO файлы

Работает в Access 2003+ (x86/x64)
Используются библиотеки FreeImage и слегка модифицированный модуль Visual Basic Wrapper for FreeImage 3 by Carsten Klein (cklein05@users.sourceforge.net) взятый с https://freeimage.sourceforge.io/download.html
Для работы требуется наличие библиотек FreeImage соответствующей разрядности.
По умолчанию ищет их в подпапке \INC\ рабочей папки примера

содержит:
PicturesFI.mdb - база Access содержащая все примеры
Pictures_Model.xml - модель объясняющая, как рассчитывается позиция текста рядом с картинкой.
\INC\ - папка содержащая библиотеки FreeImage x86/x64

Запускается с формы: ~RunMe.

![Pictures_00](https://github.com/iKaRus-VLZ/Pictures/assets/8457437/bbad6b9f-9cb2-45e6-8a3e-165387e4b3c3)

Содержит формы:

Test_PictureData_SetToControl - Пример работы с картинками
![Pictures_01](https://github.com/iKaRus-VLZ/Pictures/assets/8457437/14419edb-ec07-4e56-b213-6311203eeaa4)
Test_OlePictContinuos - Вывод картинок в ленточную форму
![Pictures_02](https://github.com/iKaRus-VLZ/Pictures/assets/8457437/9bfed7c7-d93b-4a6f-9c0f-938232e4015a)
Test_OlePictDIBContinuos - Вывод картинок встроенных в Access в ленточную форму
![Pictures_03](https://github.com/iKaRus-VLZ/Pictures/assets/8457437/3090e749-ba47-46de-a6f3-5a04dac58c62)
Sample_Clock - Аналоговые часы
![Pictures_04](https://github.com/iKaRus-VLZ/Pictures/assets/8457437/e207ee53-f88f-46a6-93db-e8dda5c6c468)
Sample_Dates - Вывод изображений в плавающие кнопки и контекстное меню
![Pictures_05](https://github.com/iKaRus-VLZ/Pictures/assets/8457437/673ada70-225f-435b-8b75-808ac36c0847)


Модуль и классы для работы с изображениями:
modPictureData - содержит функции (основные):
PictureData_SetToControl - загружает картинку в нужное свойство указанного контрола
modFreeImage - собственно модифицированный модуль Visual Basic Wrapper for FreeImage 3 by Carsten Klein (cklein05@users.sourceforge.net)
