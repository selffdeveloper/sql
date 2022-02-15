-- =============================================================================
-- KDA
-- ������ ��� ������ 76
-- exec [dbo].[elec_apteka76_dbf] @kz = 62414278, @email_in = 'd.kolesov@agrores.ru'
-- =============================================================================

create procedure [dbo].[elec_apteka76_dbf] @kz int, @email_in varchar(255) = null as


-- =============================================================================
-- ��������� ��������� �������, ���� ����� ��� ����, ����� ������� ��.
print '��������� ��������� �������, ���� ����� ��� ����, ����� ������� ��.'
if not exists (select * 
               from sysobjects 
               where name='#apteka76' and xtype='u')
			   drop table #apteka76
	
-- =============================================================================
-- ���������� ����������
print '���������� ����������'
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
       
select @date    = rtrim(ltrim(replace(convert(CHAR(10), o.[���� �������], 104), '.', '')))
	  ,@email   = c.Email_���_����������
	  ,@subject = c.����_���_����������
	  ,@doc_num = isnull(o.[����� �������], o.[��� ������])
from client as c
left join orderr as o
	on o.[��� �������] = c.[��� �������]
where o.[��� ������] = @kz

set @elec_folder  = ltrim(rtrim('\\irk\forelec\apteka76\'))
set @temlate_file = ltrim(rtrim('\\irk\forelec\apteka76\template\apteka76.dbf'))
set @out_file     = @elec_folder + ltrim(rtrim(str(@doc_num))) + '.dbf'

-- =============================================================================
-- �������
print '�������'
select rtrim(ltrim(str(convert(varchar(12), isnull(o.[����� �������], o.[��� ������]))))) as n_nacl
      ,convert(datetime, o.[���� �������]) as d_nacl
      ,rtrim(ltrim(str(convert(numeric(11,0),i.[��� ������])))) as code
      ,convert(varchar(80),dbo.str2quotestr(i.[��������])) as name
      ,convert(varchar(13),dbo.str2quotestr(case 
          when isnull(i.ean, '--') = ''
            then '--'
          else i.ean
          end)) as scancod
      ,convert(varchar(40),dbo.str2quotestr(replace(i.[�����-�������������],'|','/'))) as factory
      ,convert(varchar(20),dbo.str2quotestr(i.[������])) as country
      ,convert(numeric(14,2),rtrim(ltrim(str(oi.[����������])))) as quantity
      ,convert(numeric(14,2), dbo.calc_sebestoimost(oi.[��� ������], i.[��� ������], oi.[��� ��������], null)) as price_make
      ,convert(numeric(14,2), (oi.[����])) as price_nake
      ,convert(numeric(14,2), oi.[����] * oi.[����������]) as sum_naked
      ,convert(numeric(14,0), (isnull(oi.[������ ���] * 100, 0))) as nds_pr
      ,convert(numeric(14,2), (sum(oi.[����� ���]))) as nds_sum
      ,convert(numeric(14,2), (isnull(dbo.calc_reestr_cena(i.[��������������� ����], oi.[��� ��������], i.[��� ������]), 0))) as price_rees
      ,convert(datetime,(isnull(i.[���� �����������], ''))) as date_rees
      ,isnull(convert(char(1), i.life), '') as islife
      ,convert(varchar(15),rtrim(ltrim(isnull(ii.[���� �������], '')))) as series
      ,convert(datetime, isnull(s.[���� ��������], '')) as date_valid
      ,convert(varchar(25),dbo.str2quotestr(case 
          when isnull(ii.[����� ����������], '') = ''
            then '--'
          else ii.[����� ����������]
          end)) as gtd
      ,case when convert(varchar(70),isnull(s.[����� �����������], '')) = '' then '--' else convert(varchar(70),ltrim(rtrim(s.[����� �����������]))) end as sert
      ,convert(varchar(30),'�����������') as filial
      ,convert(varchar(30),c.[��� �������]) /*c.[��������]*/ as apteka
into #apteka76
from out_item as oi
left join orderr as o
	on o.[��� ������] = oi.[��� ������]
left join client as c
	on c.[��� �������] = o.[��� �������]
left join info as i
	on i.[��� ������] = oi.[��� ������]
left join in_item as ii
	on ii.[��� ������] = oi.[��� ������] and ii.[��� ��������] = oi.[��� ��������]
left join series as s
	on s.[��� ������] = ii.[��� ������] and s.[���� �������] = ii.[���� �������]
where o.[��� ������] = @kz
GROUP BY o.[����� �������],o.[��� ������],o.[���� �������],i.[��� ������],i.��������,i.[�����-�������������],i.������
,oi.����������,oi.����,oi.[������ ���],s.[���� ��������],ii.[����� ����������],ii.[���� �������],s.[���� ������ �����������]
,s.center_sert,s.[����� �����������],ii.[���� �������],i.ean,i.[��������������� ����],oi.[����� � ���],i.life,c.[��������]
,c.[��� �������],i.[���� �����������],oi.[��� ������],oi.[��� ��������],ii.����������,i.ean13
ORDER BY NAME

-- =============================================================================
-- ���������� ������
print '���������� ������'
declare @copy_to varchar(256) = '\\irk\forelec\apteka76\apteka76.dbf'
exec copy_file @from = @temlate_file, @to = @copy_to

-- =============================================================================
-- ��������� ������ � ����� ������
print '��������� ������ � ����� ������'
set @sql_str = 'insert into opendatasource(''microsoft.ace.oledb.12.0'',''data source = ' + ltrim(rtrim(@elec_folder)) + ' ;extended properties = dbase iv'')...apteka76 select * from #apteka76'
    begin try
        execute(@sql_str)
    end try
    begin catch
		print '================================================================='
		print '������:'
		print error_message()
		print '================================================================='
		print char(1)
		print '================================================================='
		print '������:'
		print @sql_str
		print '================================================================='
    end catch
	
-- =============================================================================
-- ��������������� ������

print '��������������� ������'
declare @old_file_name varchar(256) = ltrim(rtrim(@elec_folder)) + 'apteka76.dbf' 
declare @new_file_name varchar(256) = ltrim(rtrim(@doc_num)) + '.dbf'
exec dbo.rename_file @file_name_in = @old_file_name, @file_name_out = @new_file_name

-- =============================================================================
-- ���������� ���� �� �����
print '���������� ���� �� �����'

if @email_in is not null set @email = @email_in

exec send_mail
	 @email = @email,
	 @subject = @subject,
	 @file_path = @out_file,
	 @message = null
	 
-- =============================================================================
-- ���������� ������ � sended
print '���������� ������ � sended'
declare @move_to   varchar(256) = @elec_folder + ltrim(rtrim('\sended\')) 
exec move_file @from = @out_file, @to = @move_to

drop table #apteka76