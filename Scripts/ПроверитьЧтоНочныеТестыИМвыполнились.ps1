$folder = "\\WorkServer\ТестыИМ\Тесты\"
#$folder = "\\WorkServer\Share\Admin1C\Logs\"

cd $folder

$files = dir лог_тестирования.* | where {$_.LastWriteTime -lt (date).Date} | select -Property Name, LastWriteTime 
#$files = dir *IM_test*.txt | where {$_.LastWriteTime -lt (date).Date} | select -Property Name, LastWriteTime 
if ($files -ne $null)
{
    cls #чтобы не показывалась ошибка иницилизации провайдера
   "Список непрошедших тестов для ИМ Тест 1С - не отработали сегодня ночью"
   $files
   pause
}
