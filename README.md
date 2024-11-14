<h1 align="center">Инструкция по автоматическому созданию ссылок и закладок в PDF документе</h1>

  <details>
    <p>
    <summary>
      <b>1</b> Настройка вкладок
    </summary>
    </p>
    <details>
      <p>
      <summary>
        <b>1.1</b> Создание нового стиля
      </summary>
      </p>
      <p>
      <b>1.1.1</b> Первым делом создаём адекватный стиль заголовка. Переходим во вкладку "Главная". В разделе "Стили" снизу справа нажимаем зночок <img src="https://github.com/Mr-Krabs95/links_and_bookmarks_PDF/blob/adding_screenshots/screenshots/icons/icon1.png"> ->"Создать стиль".<br>
      <img src="https://github.com/Mr-Krabs95/links_and_bookmarks_PDF/blob/adding_screenshots/screenshots/1.1.1.png">
      </p>
      </p>
      <b>1.1.2</b> В появившемся окне "Создание стиля" вводим следующие значения: <br>
      Имя: Закладки PDF<br>
      Стиль: Абзаца<br>
      Основан на стиле: Обычный<br>
      Стиль следующего абзаца: Обычный<br>
      Форматирование: Times New Roman, 12, <b>Ж</b>, Авто, выравнивание по ширине, междустрочный интервал - одинарный, междустрочное расстояние - минимальное<br>
      Ставим галочку "Добавить в коллекцию стилей"<br>
      <img src="https://github.com/Mr-Krabs95/links_and_bookmarks_PDF/blob/adding_screenshots/screenshots/1.2.png">
      </p>
      </p>
      <b>1.1.3</b> Далее нажимаем "Формат" -> "Абзац..." <img src="https://github.com/Mr-Krabs95/links_and_bookmarks_PDF/blob/adding_screenshots/screenshots/icons/icon2.png"><br>
      В открывшемся окне следующие настройки:<br>
      Уровень: Уровень 1<br>
      Отступы первая строка: отступ на 1.25 см и нажимаем OK -> OK.<br>
      <img src="https://github.com/Mr-Krabs95/links_and_bookmarks_PDF/blob/adding_screenshots/screenshots/1.3.png"><br>
      </p>
    </details>
    <details>
      <p>
      <summary>
        <b>1.2</b> Создание заголовков
      </summary>
      </p>
      <p>
      <b>1.2</b> Выделите текст, который хотите отформатировать. Во вкладке "Главная" в разделе "Стили" нажмите на созданный только что стиль. Пробегитесь по текстовой части и примените стиль ко всем заголовкам, которые нужно будет отображать во вкладках PDF документа и на которые будем делать ссылки в дальнейшем.<br>
      <img src="https://github.com/Mr-Krabs95/links_and_bookmarks_PDF/blob/adding_screenshots/screenshots/1.4.png"><br>
      </p>
      <p>
      <b>1.5</b> Для удобной навигации в по документу Word во вкладке "Вид" в разделе "Отображение" нажмите галочку "Область навигации". Слева появится панель "Навигация". Проверьте ваши будущие вкладки и перейдите к "Состав проектной документации", нажав на соответствующий заголовок на панели навигации, для дальнейшей настройки.<br>
      <img src="https://github.com/Mr-Krabs95/links_and_bookmarks_PDF/blob/adding_screenshots/screenshots/1.5.png"><br>
      </p>
    </details>
  </details>
  <details>
    <p>
    <summary>
      <b>2</b> Настройка ссылок
    </summary>
    </p>
    <details>
      <p>
      <summary>
        <b>2.1</b> Добавление вкладок
      </summary>
      </p>
      <p>
      <b>2.1.1</b> Выделите заголовок и перейдите во вкладку "Вставка". В разделе "Ссылки" нажмите "Закладка".<br>
      <img src="https://github.com/Mr-Krabs95/links_and_bookmarks_PDF/blob/adding_screenshots/screenshots/2.1.1.png">
      </p>
      <p>
      <b>2.1.2</b> В открывшемся окне введите имя закладки. Оно не должно начинаться с цифры, содержать пробелов и каких-либо символов кроме нижнего подчёркивания _ . Нажмите "Добавить".<br>
      <img src="https://github.com/Mr-Krabs95/links_and_bookmarks_PDF/blob/adding_screenshots/screenshots/2.1.2.png">
      </p>
      <p>
      <b>2.1.3</b> Продолжайте повторять процедуру для всех заголовков, на которые в дальнейшем мы будем давать ссылки в содержании.<br>
      Для автоматизации процесса можно написать макрос. Как это сделать описано в п. 2.2.<br>
      </p>
    </details>
    <details>
      <p>
      <summary>
        <b>2.2</b> Создание макроса (по желанию)
      </summary>
      </p>
      <p>
      <b>2.2.1</b> Нажмите сочетание клавиш Alt+F11. В появившемся окне "Insert" -> "Module".
      <img src="https://github.com/Mr-Krabs95/links_and_bookmarks_PDF/blob/adding_screenshots/screenshots/2.2.1.png">
      </p>
      <p>
      <b>2.2.2</b> В открывшемся окне вставьте следующий код:
      </p>

            Sub AddBookmarksToCustomHeadings()
            Dim para As Paragraph
            Dim bookmarkName As String
            Dim textExcerpt As String
            Dim i As Integer
        
            i = 1 ' Счётчик для уникальных имён, если заголовок повторяется
        
            For Each para In ActiveDocument.Paragraphs
                ' Проверка, является ли стиль абзаца пользовательским заголовком
                If para.Style = "Закладки PDF" Then
                    ' Получаем первые 50 символов текста абзаца
                    textExcerpt = Left(para.Range.Text, 50)
                    
                    ' Удаляем цифры и пробелы в начале строки, если они есть
                    If IsNumeric(Left(textExcerpt, 1)) Then
                        textExcerpt = Trim(Mid(textExcerpt, InStr(1, textExcerpt, " ") + 1))
                    End If
                    
                    ' Заменяем пробелы и недопустимые символы
                    textExcerpt = Replace(textExcerpt, " ", "_")
                    textExcerpt = Replace(textExcerpt, vbTab, "_")
                    textExcerpt = Replace(textExcerpt, ".", "")
                    textExcerpt = Replace(textExcerpt, ",", "")
                    textExcerpt = Replace(textExcerpt, ":", "")
                    textExcerpt = Replace(textExcerpt, ";", "")
                    textExcerpt = Replace(textExcerpt, "!", "")
                    textExcerpt = Replace(textExcerpt, "?", "")
                    textExcerpt = Replace(textExcerpt, "\", "")
                    textExcerpt = Replace(textExcerpt, "/", "")
                    textExcerpt = Replace(textExcerpt, "[", "")
                    textExcerpt = Replace(textExcerpt, "]", "")
                    textExcerpt = Replace(textExcerpt, "(", "")
                    textExcerpt = Replace(textExcerpt, ")", "")
                    textExcerpt = Replace(textExcerpt, "'", "")
                    textExcerpt = Replace(textExcerpt, """", "")
                    
                    ' Проверка на существование закладки и создание уникального имени
                    bookmarkName = textExcerpt & "_" & i
        
                    ' Проверка, если закладка с таким именем уже существует, удаляем её
                    If ActiveDocument.Bookmarks.Exists(bookmarkName) Then
                        ActiveDocument.Bookmarks(bookmarkName).Delete
                    End If
        
                    ' Добавляем новую закладку на абзац
                    On Error Resume Next ' В случае ошибки (например, имя закладки всё ещё некорректно)
                    ActiveDocument.Bookmarks.Add Range:=para.Range, Name:=bookmarkName
                    On Error GoTo 0 ' Отключаем обработку ошибок
        
                    ' Увеличиваем счётчик
                    i = i + 1
                End If
            Next para
            
            MsgBox "Закладки добавлены к заголовкам."
        End Sub

  Для выполнения нажмите F5.
  </details>
  <details>
      <p>
      <summary>
        <b>2.3</b> Добавление ссылок
      </summary>
      </p>
      <p>
      <b>2.3.1</b> Перед тем, как вставлять ссылки нужно отменить их автоматическое форматирование и изменение цвета после нажатия. Для этого перейдите во вкладку "Главное". В разделе "Стили" нажмите значок <img src="https://github.com/Mr-Krabs95/links_and_bookmarks_PDF/blob/adding_screenshots/screenshots/icons/icon1.png"><br>
      В открывшемся справа окне найдите стиль "Гиперссылка" -> Правая конпка мыши -> "Изменить..."<br>
      <img src="https://github.com/Mr-Krabs95/links_and_bookmarks_PDF/blob/adding_screenshots/screenshots/2.3.1.png">
      </p>
      <p>
      <b>2.3.2</b> Настраиваем параметры форматирования: Шрифт: Times New Roman, 12 пт, без подчеркивания, Цвет шрифта: Авто. Нажмите OK.
      </p>
      <p>
      <b>2.3.3</b> Как сделать так, чтобы ссылка не подчёркивалась после нажатия я так и не нашёл. Кто занет - напишите мне =)<br>
      А чтобы ссылка не изменялась после использования - не нажимайте на неё в Word.
      </p>
      <p>
      <b>2.3.3</b> Для добавления ссылки на вкладку, созданную в п. 2.1, выделите текст в содержании тома. Затем перейдите во вкладку "Вставка". В разделе "Ссылки" нажмите "Ссылка" <img src="https://github.com/Mr-Krabs95/links_and_bookmarks_PDF/blob/adding_screenshots/screenshots/icons/icon3.png" height=50 width=75><br>
      <img src="https://github.com/Mr-Krabs95/links_and_bookmarks_PDF/blob/adding_screenshots/screenshots/2.3.3.png">
      </p>
      <p>
      <b>2.3.4</b> В открывшемся окне в разделе "Связать с:" выберите "Место в документе". Отобразятся вкладки, которые мы добавили ранее в п. 2.1-2.2.<br>
      Продолжайте добавлять ссылки в содержании для всей текстовой части и переходите к п. 3.
      </p>
  </details>
</details>
<details>
  <p>
  <summary>
    <b>3</b> Создание PDF документа
  </summary>
    <p>
    <b>3.1</b> Для конвертации DOC в PDF необходимо перейти во вкладку "Файл". В разделе "Сохранить как" нажмите "Обзор". В появившемся окне укажите директорию, имя файла и тип файла PDF<br>
    <img src="https://github.com/Mr-Krabs95/links_and_bookmarks_PDF/blob/adding_screenshots/screenshots/3.1.png">
    </p>
    <p>
    <b>3.2</b> Нажмите кнопку "Параметры". В появившемся окне в разделе "Включить непечатаемые данные" поставьте галочку "Создать закладки, используя:" выберите пункт "заголовки". Нажмите OK.<br>
    <img src="https://github.com/Mr-Krabs95/links_and_bookmarks_PDF/blob/adding_screenshots/screenshots/3.2.png">
    </p>
    <p>
    <b>3.3</b> Нажмите кнопку "Сохранить". Зайдите в директорию выбранную при сохранении, откройфе PDF документ и наслаждайтесь результатом АВТОМАТИЗАЦИИ.<br>
    <img src="https://github.com/Mr-Krabs95/links_and_bookmarks_PDF/blob/adding_screenshots/screenshots/3.3.png">
    </p>
</details>
