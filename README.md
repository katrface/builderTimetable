# Построить удобное расписание за 19 кликов

VBA Excel макрос и пользовательская форма для построения расписания определенной группы на основе excel файла опубликованного КФУ ИВМиИТ.
Может использоваться для других институтов с незначительными изменениями.

## Пример работы макроса
![Пример работы макроса](https://sun9-53.userapi.com/WcjFrKTpsIwcXjQZPF8sJGbPfJVJ3BrOlMEG7Q/cG6gD3evUF4.jpg)

## Порядок действий:
### Подготовить Excel
1. Изменить формат на xlsm
В открытом excel файле с расписанием:

   **Файл** - **Сохранить как** - Тип файла: **Книга Excel с поддержкой макросов (*.xlsm)** - **Сохранить**

![Тип файла xlsm](https://sun9-73.userapi.com/7YMsKYm6vunjmyygCdfCnx4mDxEGfTBUsssBaA/eRm7Kuv_49Q.jpg)

2. Включить вкладку "Разработчик"

   **Файл** - **Параметры** - **Настройка ленты** - В списке **Основные вкладки** установите флажок **Разработчик** - **ОК**

![Включение вкладки "Разработчик"](https://sun3-11.userapi.com/85KvEttZUjKtgCggDTUQIpn6kMAw8l97NagRNw/GRJiG5_LN-U.jpg)

3. Импортировать макрос и пользовательскую форму

   Вкладка **Разработчик** - **Visual Basic** - Перенесите файлы в окно структуры VBA Project (или импортируйте другим способом) - Закройте VBA

![Подключение файлов к проекту VBA](https://sun9-20.userapi.com/NNed5_rZ9QuadW3hLyjmazv0t8dUodf6VcuZrQ/KYMSshoagRY.jpg)
   
> При подключении файлов может возникнуть ошибка связанная с файлом c расширением .frx - Нажмите **ОК** - Переходите к следующему этапу


### Вызвать пользовательскую форму
Вкладка **Разработчик** - **Макросы** - Выбираем **showUserForm** - **Выполнить**

![Вызов макроса showUserForm](https://sun9-47.userapi.com/FVtgGLmt864CBUXn3F3xpXNpYovfPPoiquADEA/SzOUYi4gFjE.jpg)
