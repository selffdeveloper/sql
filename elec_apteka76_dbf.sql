-- =============================================================================
-- KDA
-- Формат для Аптеки 76
-- exec [dbo].[elec_apteka76_dbf] @kz = 62414278, @email_in = 'd.kolesov@agrores.ru'
-- =============================================================================

create procedure [dbo].[elec_apteka76_dbf] @kz int, @email_in varchar(255) = null as


-- =============================================================================
-- Проверяем временную таблицу, если такая уже есть, тогда дропаем ее.
print 'Проверяем временную таблицу, если такая уже есть, тогда дропаем ее.'
if not exists (select * 
               from sysobjects 
               where name='#apteka76' and xtype='u')
			   drop table #apteka76
	
-- =============================================================================
-- Объявление переменных
print 'Объявление переменных'
declare @cmd              varchar(512)
       ,@result           int
       ,@elec_folder      varchar(256)
       ,@temlate_file     varchar(256)
       ,@email			  varchar(256)
       ,@subject          varchar(256)
       ,@date             char(8)
       ,@doc_num		  int
       ,@out_file		  varchar(256)
       ,@sql_str		  varchar(1024)
       
select @date    = rtrim(ltrim(replace(convert(CHAR(10), o.[Дата фактуры], 104), '.', '')))
	  ,@email   = c.Email_для_электронки
	  ,@subject = c.Тема_для_электронки
	  ,@doc_num = isnull(o.[Номер фактуры], o.[Код заказа])
from client as c
left join orderr as o
	on o.[Код клиента] = c.[Код клиента]
where o.[Код заказа] = @kz

set @elec_folder  = ltrim(rtrim('\\irk\forelec\apteka76\'))
set @temlate_file = ltrim(rtrim('\\irk\forelec\apteka76\template\apteka76.dbf'))
set @out_file     = @elec_folder + ltrim(rtrim(str(@doc_num))) + '.dbf'

-- =============================================================================
-- Выборка
print 'Выборка'
select rtrim(ltrim(str(convert(varchar(12), isnull(o.[номер фактуры], o.[код заказа]))))) as n_nacl
      ,convert(datetime, o.[дата фактуры]) as d_nacl
      ,rtrim(ltrim(str(convert(numeric(11,0),i.[код товара])))) as code
      ,convert(varchar(80),dbo.str2quotestr(i.[название])) as name
      ,convert(varchar(13),dbo.str2quotestr(case 
          when isnull(i.ean, '--') = ''
            then '--'
          else i.ean
          end)) as scancod
      ,convert(varchar(40),dbo.str2quotestr(replace(i.[фирма-производитель],'|','/'))) as factory
      ,convert(varchar(20),dbo.str2quotestr(i.[страна])) as country
      ,convert(numeric(14,2),rtrim(ltrim(str(oi.[количество])))) as quantity
      ,convert(numeric(14,2), dbo.calc_sebestoimost(oi.[код заказа], i.[код товара], oi.[код поставки], null)) as price_make
      ,convert(numeric(14,2), (oi.[цена])) as price_nake
      ,convert(numeric(14,2), oi.[цена] * oi.[количество]) as sum_naked
      ,convert(numeric(14,0), (isnull(oi.[ставка ндс] * 100, 0))) as nds_pr
      ,convert(numeric(14,2), (sum(oi.[сумма ндс]))) as nds_sum
      ,convert(numeric(14,2), (isnull(dbo.calc_reestr_cena(i.[ориентировочная цена], oi.[код поставки], i.[код товара]), 0))) as price_rees
      ,convert(datetime,(isnull(i.[дата регистрации], ''))) as date_rees
      ,isnull(convert(char(1), i.life), '') as islife
      ,convert(varchar(15),rtrim(ltrim(isnull(ii.[дата выпуска], '')))) as series
      ,convert(datetime, isnull(s.[срок годности], '')) as date_valid
      ,convert(varchar(25),dbo.str2quotestr(case 
          when isnull(ii.[номер декларации], '') = ''
            then '--'
          else ii.[номер декларации]
          end)) as gtd
      ,case when convert(varchar(70),isnull(s.[номер сертификата], '')) = '' then '--' else convert(varchar(70),ltrim(rtrim(s.[номер сертификата]))) end as sert
      ,convert(varchar(30),'агроресурсы') as filial
      ,convert(varchar(30),c.[код клиента]) /*c.[название]*/ as apteka
into #apteka76
from out_item as oi
left join orderr as o
	on o.[Код заказа] = oi.[Код заказа]
left join client as c
	on c.[Код клиента] = o.[Код клиента]
left join info as i
	on i.[Код товара] = oi.[Код товара]
left join in_item as ii
	on ii.[Код товара] = oi.[Код товара] and ii.[Код поставки] = oi.[Код поставки]
left join series as s
	on s.[Код товара] = ii.[Код товара] and s.[Дата выпуска] = ii.[Дата выпуска]
where o.[Код заказа] = @kz
GROUP BY o.[Номер фактуры],o.[Код заказа],o.[Дата фактуры],i.[Код товара],i.Название,i.[Фирма-производитель],i.Страна
,oi.Количество,oi.Цена,oi.[Ставка НДС],s.[Срок годности],ii.[Номер декларации],ii.[Дата выпуска],s.[Дата выдачи сертификата]
,s.center_sert,s.[Номер сертификата],ii.[Дата выпуска],i.ean,i.[Ориентировочная цена],oi.[Сумма с НДС],i.life,c.[Название]
,c.[Код клиента],i.[дата регистрации],oi.[Код заказа],oi.[Код поставки],ii.маркировка,i.ean13
ORDER BY NAME

-- =============================================================================
-- Перемещаем шаблон
print 'Перемещаем шаблон'
declare @copy_to varchar(256) = '\\irk\forelec\apteka76\apteka76.dbf'
exec copy_file @from = @temlate_file, @to = @copy_to

-- =============================================================================
-- Заполняем шаблон и ловим ошибки
print 'Заполняем шаблон и ловим ошибки'
set @sql_str = 'insert into opendatasource(''microsoft.ace.oledb.12.0'',''data source = ' + ltrim(rtrim(@elec_folder)) + ' ;extended properties = dbase iv'')...apteka76 select * from #apteka76'
    begin try
        execute(@sql_str)
    end try
    begin catch
		print '================================================================='
		print 'Ошибка:'
		print error_message()
		print '================================================================='
		print char(1)
		print '================================================================='
		print 'Запрос:'
		print @sql_str
		print '================================================================='
    end catch
	
-- =============================================================================
-- Переименовываем файлик

print 'Переименовываем файлик'
declare @old_file_name varchar(256) = ltrim(rtrim(@elec_folder)) + 'apteka76.dbf' 
declare @new_file_name varchar(256) = ltrim(rtrim(@doc_num)) + '.dbf'
exec dbo.rename_file @file_name_in = @old_file_name, @file_name_out = @new_file_name

-- =============================================================================
-- Отправляем файл на почту
print 'Отправляем файл на почту'

if @email_in is not null set @email = @email_in

exec send_mail
	 @email = @email,
	 @subject = @subject,
	 @file_path = @out_file,
	 @message = null
	 
-- =============================================================================
-- Перемещаем файлик в sended
print 'Перемещаем файлик в sended'
declare @move_to   varchar(256) = @elec_folder + ltrim(rtrim('\sended\')) 
exec move_file @from = @out_file, @to = @move_to

drop table #apteka76