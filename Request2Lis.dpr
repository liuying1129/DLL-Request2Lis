library Request2Lis;

{ Important note about DLL memory management: ShareMem must be the
  first unit in your library's USES clause AND your project's (select
  Project-View Source) USES clause if your DLL exports any procedures or
  functions that pass strings as parameters or function results. This
  applies to all strings passed to and from your DLL--even those that
  are nested in records and classes. ShareMem is the interface unit to
  the BORLNDMM.DLL shared memory manager, which must be deployed along
  with your DLL. To avoid using BORLNDMM.DLL, pass string information
  using PChar or ShortString parameters. }

//==============================================================================
//��DLL�ṩ���������ڽ�JSON��ʽ�ļ���������Ϣ����LIS
//3���������:
//AAdoconnstr:LIS���ݿ�������ַ���
//ARequestJSON:JSON��ʽ�ļ������뵥��Ϣ
//CurrentWorkGroup:��ǰ������.�������Ŀδ����Ĭ�Ϲ�����,���뵥�����뵱ǰ������
//JSONʵ��:
//  {
//      "JSON����Դ":"HIS",
//      "����ҽ��": [
//          {
//              "���뵥���": "10000",
//              "������": "101234",
//              "��������": "�ܲ�",
//              "�����Ա�": "��",
//              "��������": "24��",
//              "��������": "2023-2-6",
//              "�������": "�ڿ�",
//              "����ҽ��": "���",
//              "����":"",
//              "�ٴ����":"",
//              "��ע":"",
//              "������˾":"",
//              "��������":"",
//              "����":"",
//              "����":"",
//              "���":"",
//              "����":"",
//              "סַ":"",
//              "�绰":"",
//              "�ⲿϵͳΨһ���":"",
//              "ҽ����ϸ": [
//                  {
//                      "������": "S0087",
//                      "LIS�����Ŀ����": "06",
//                      "�����": "12345",
//                      "���ȼ���": "����",
//                      "��������": "Ѫ��",
//                      "����״̬": "����"
//                  },
//                  {
//                      "������": "X0013",
//                      "LIS�����Ŀ����": "54",
//                      "�����": "12346",
//                      "���ȼ���": "����",
//                      "��������": "ȫѪ",
//                      "����״̬": "����"
//                  }
//              ]
//          },
//          {
//              "���뵥���": "10001",
//              "������": "101221",
//              "��������": "����",
//              "�����Ա�": "��",
//              "��������": "25��",
//              "��������": "2023-2-7",
//              "�������": "���",
//              "����ҽ��": "�ŷ�",
//              "����":"",
//              "�ٴ����":"",
//              "��ע":"",
//              "������˾":"",
//              "��������":"",
//              "����":"",
//              "����":"",
//              "���":"",
//              "����":"",
//              "סַ":"",
//              "�绰":"",
//              "�ⲿϵͳΨһ���":"",
//              "ҽ����ϸ": [
//                  {
//                      "������": "S0088",
//                      "LIS�����Ŀ����": "06",
//                      "�����": "12347",
//                      "���ȼ���": "����",
//                      "��������": "Ѫ��",
//                      "����״̬": "����"
//                  },
//                  {
//                      "������": "X0014",
//                      "LIS�����Ŀ����": "54",
//                      "�����": "12348",
//                      "���ȼ���": "����",
//                      "��������": "ȫѪ",
//                      "����״̬": "����"
//                  }
//              ]
//          }
//      ]
//  }
//JSON������ڵ�key��JSON����Դ������ҽ���������(�ر�ģ����JSON����Դ��ֵΪExcel����key���Բ�����)��ҽ����ϸ
//��JSON����Դ��ֵ���HIS��Excel
//������š�������JSON����Դ��ֵΪHISʱ,������š��ǳ���������Ŀ����ͬһ�ż��鵥���ж�����
//���ⲿϵͳΨһ��š�: ����JSON����Դ��ֵΪHISʱ,HIS/PEIS���ⲿϵͳ���ô˱�Ź����ܼ����������.�˱���п���������,Ҳ�п�����HIS��ʾ�˴ο����Ŀ�����
//�����LIS�����Ŀ���롿��ֵ��LIS�в����ڣ�����ᵼ�벡�˻�����Ϣ�����ᵼ�������Ŀ
//���ϣ�������벡�˻�����Ϣ,����Ҫ��֤��ҽ����ϸ��������һ����¼,������һ����Ч���ݵļ�¼
//JSON������ʱ���ʽ��YYYY-MM-DD hh:nn:ss
//
//2023-02-17�������Ѹ��ݹ����顢��������Ϊ���ݽ��в�
//�Ƿ�Ҫ��������Ŀ��������ĸ�����в𵥣��۲�Ӧ������ٶ�
//
//��JSON��ʾΪ��ͼ����վ:https://jsoncrack.com/editor
//==============================================================================

uses
  SysUtils,
  Classes,
  ADODB,
  DateUtils,
  Dialogs,
  superobject in 'superobject.pas';

{$R *.res}
function GetServerDate(AConnectionString:string): TDateTime;
var
  Conn:TADOConnection;
  adotempDate:tadoquery;
begin
  Conn:=TADOConnection.Create(nil);
  Conn.LoginPrompt:=false;
  Conn.ConnectionString:=AConnectionString;
  adotempDate:=tadoquery.Create(NIL);
  ADOTEMPDATE.Connection:=Conn;
  ADOTEMPDATE.Close;
  ADOTEMPDATE.SQL.Clear;
  ADOTEMPDATE.SQL.Text:='SELECT GETDATE() as ServerDate ';
  ADOTEMPDATE.Open;
  result:=ADOTEMPDATE.fieldbyname('ServerDate').AsDateTime;
  ADOTEMPDATE.Free;
  Conn.Free;
end;

function ExecSQLCmd(AConnectionString:string;ASQL:string):integer;
var
  Conn:TADOConnection;
  Qry:TAdoQuery;
begin
  Conn:=TADOConnection.Create(nil);
  Conn.LoginPrompt:=false;
  Conn.ConnectionString:=AConnectionString;
  Qry:=TAdoQuery.Create(nil);
  Qry.Connection:=Conn;
  Qry.Close;
  Qry.SQL.Clear;
  Qry.SQL.Text:=ASQL;
  Try
    Result:=Qry.ExecSQL;
  except
    on E:Exception do
    begin
      MessageDlg('����ExecSQLCmdʧ��:'+E.Message+'�������SQL:'+ASQL,mtError,[MBOK],0);
      Result:=-1;
    end;
  end;
  Qry.Free;
  Conn.Free;
end;

function ScalarSQLCmd(AConnectionString:string;ASQL:string):string;
var
  Conn:TADOConnection;
  Qry:TAdoQuery;
begin
  Result:='';
  Conn:=TADOConnection.Create(nil);
  Conn.LoginPrompt:=false;
  Conn.ConnectionString:=AConnectionString;
  Qry:=TAdoQuery.Create(nil);
  Qry.Connection:=Conn;
  Qry.Close;
  Qry.SQL.Clear;
  Qry.SQL.Text:=ASQL;
  Try
    Qry.Open;
  except
    on E:Exception do
    begin
      MessageDlg('����ScalarSQLCmdʧ��:'+E.Message+'�������SQL:'+ASQL,mtError,[MBOK],0);
      Qry.Free;
      Conn.Free;
      exit;
    end;
  end;
  Result:=Qry.Fields[0].AsString;
  Qry.Free;
  Conn.Free;
end;

//��ҽ��JSON������LIS
procedure RequestForm2Lis(const AAdoconnstr,ARequestJSON,CurrentWorkGroup:PChar);stdcall;
var
  adoconn11,adoconn22:Tadoconnection;
  adotemp11,adotemp22:Tadoquery;
  aJson:ISuperObject;
  aSuperArray,aSuperArrayMX: TSuperArray;
  i,j:integer;
  defaultWorkGroup:string;//Ĭ�Ϲ�����
  defaultSampleType:string;//Ĭ����������
  WorkGroup:string;
  SampleType:string;
  chk_con_unid:string;
  YXJB:STRING;//���ȼ���
  SampleStatus:string;//����״̬
  fs:TFormatSettings;
  RequestDate:TDateTime;//��������
  ServerDateTime:TDateTime;
  lsh:string;//��ˮ��
  pkcombin_id:String;//LIS�����Ŀ����
  RequestDateStr:String;//��������
  checkid:String;//������
  patientname:String;//��������
  sex:String;//�����Ա�
  age:String;//��������
  Caseno:String;//������
  deptname:String;//�������
  check_doctor:String;//����ҽ��
  bedno:String;//����
  diagnose:String;//�ٴ����
  issure:String;//��ע
  WorkCompany:String;//������˾
  WorkDepartment:String;//��������
  WorkCategory:String;//����
  WorkID:String;//����
  ifMarry:String;//���
  OldAddress:String;//����
  Address:String;//סַ
  Telephone:String;//�绰
  DNH:String;//���뵥���(HIS)
  His_Unid:String;//�ⲿϵͳΨһ���(HIS)
begin
  ServerDateTime:=GetServerDate(AAdoconnstr);

  aJson:=SO(ARequestJSON);
  if not aJson.AsObject.Exists('JSON����Դ') then exit;//�ж�key�Ƿ���ڵ���һ��д��:if aJson['JSON����Դ']=nil then exit;
  if not aJson.AsObject.Exists('����ҽ��') then exit;
  
  aSuperArray:=aJson['����ҽ��'].AsArray;
  for i:=0 to aSuperArray.Length-1 do
  begin
    if not aSuperArray[i].AsObject.Exists('ҽ����ϸ') then continue;

    aSuperArrayMX:=aSuperArray[i]['ҽ����ϸ'].AsArray;
    for j:=0 to aSuperArrayMX.Length-1 do
    begin
      if ('Excel'<>aJson.S['JSON����Դ'])and(not aSuperArrayMX[j].AsObject.Exists('�����')) then continue;

      if aSuperArrayMX[j].AsObject.Exists('LIS�����Ŀ����') then pkcombin_id:=aSuperArrayMX[j]['LIS�����Ŀ����'].AsString else pkcombin_id:=''; 
      if pkcombin_id='' then pkcombin_id:='�����ڵ������Ŀ����';
        
      defaultWorkGroup:=ScalarSQLCmd(AAdoconnstr,'select dept_DfValue from combinitem where Id='''+pkcombin_id+''' ');
      defaultSampleType:=ScalarSQLCmd(AAdoconnstr,'select specimentype_DfValue from combinitem where Id='''+pkcombin_id+''' ');

      //���Ĭ�Ϲ�����Ϊ��,���뵱ǰ������
      WorkGroup:=defaultWorkGroup;
      if WorkGroup='' then WorkGroup:=CurrentWorkGroup;
      if 'Excel'=aJson['JSON����Դ'].AsString then WorkGroup:=CurrentWorkGroup;

      //���JSON����������Ϊ��,��ȡĬ����������
      if aSuperArrayMX[j].AsObject.Exists('��������') then SampleType:=aSuperArrayMX[j]['��������'].AsString else SampleType:=''; 
      if (SampleType='')and('Excel'<>aJson['JSON����Դ'].AsString) then SampleType:=defaultSampleType;

      if aSuperArrayMX[j].AsObject.Exists('���ȼ���') then YXJB:=aSuperArrayMX[j]['���ȼ���'].AsString else YXJB:=''; 
      if YXJB='' then YXJB:='����';

      if aSuperArrayMX[j].AsObject.Exists('����״̬') then SampleStatus:=aSuperArrayMX[j]['����״̬'].AsString else SampleStatus:=''; 
      if SampleStatus='' then SampleStatus:='����';

      fs.DateSeparator:='-';
      fs.TimeSeparator:=':';
      fs.ShortDateFormat:='YYYY-MM-DD hh:nn:ss';
      if aSuperArray[i].AsObject.Exists('��������') then RequestDateStr:=aSuperArray[i]['��������'].AsString else RequestDateStr:='';
      RequestDate:=StrtoDateTimeDef(RequestDateStr,ServerDateTime,fs);
      if  RequestDate<2 then ReplaceDate(RequestDate,ServerDateTime);//��ʾ1899-12-30,û�и����ڸ�ֵ
      if (HourOf(RequestDate)=0) and (MinuteOf(RequestDate)=0) and (SecondOf(RequestDate)=0) then ReplaceTime(RequestDate,ServerDateTime);//��ʾû�и�ʱ�丳ֵ

      if aSuperArrayMX[j].AsObject.Exists('������') then checkid:=aSuperArrayMX[j]['������'].AsString else checkid:='';
      if aSuperArray[i].AsObject.Exists('��������') then patientname:=aSuperArray[i]['��������'].AsString else patientname:='';
      if aSuperArray[i].AsObject.Exists('�����Ա�') then sex:=aSuperArray[i]['�����Ա�'].AsString else sex:=''; 
      if aSuperArray[i].AsObject.Exists('��������') then age:=aSuperArray[i]['��������'].AsString else age:='';
      if aSuperArray[i].AsObject.Exists('������') then Caseno:=aSuperArray[i]['������'].AsString else Caseno:='';
      if aSuperArray[i].AsObject.Exists('�������') then deptname:=aSuperArray[i]['�������'].AsString else deptname:='';
      if aSuperArray[i].AsObject.Exists('����ҽ��') then check_doctor:=aSuperArray[i]['����ҽ��'].AsString else check_doctor:='';
      if aSuperArray[i].AsObject.Exists('����') then bedno:=aSuperArray[i]['����'].AsString else bedno:='';
      if aSuperArray[i].AsObject.Exists('�ٴ����') then diagnose:=aSuperArray[i]['�ٴ����'].AsString else diagnose:='';
      if aSuperArray[i].AsObject.Exists('��ע') then issure:=aSuperArray[i]['��ע'].AsString else issure:='';
      if aSuperArray[i].AsObject.Exists('������˾') then WorkCompany:=aSuperArray[i]['������˾'].AsString else WorkCompany:='';
      if aSuperArray[i].AsObject.Exists('��������') then WorkDepartment:=aSuperArray[i]['��������'].AsString else WorkDepartment:='';
      if aSuperArray[i].AsObject.Exists('����') then WorkCategory:=aSuperArray[i]['����'].AsString else WorkCategory:='';
      if aSuperArray[i].AsObject.Exists('����') then WorkID:=aSuperArray[i]['����'].AsString else WorkID:='';
      if aSuperArray[i].AsObject.Exists('���') then ifMarry:=aSuperArray[i]['���'].AsString else ifMarry:='';
      if aSuperArray[i].AsObject.Exists('����') then OldAddress:=aSuperArray[i]['����'].AsString else OldAddress:='';
      if aSuperArray[i].AsObject.Exists('סַ') then Address:=aSuperArray[i]['סַ'].AsString else Address:='';
      if aSuperArray[i].AsObject.Exists('�绰') then Telephone:=aSuperArray[i]['�绰'].AsString else Telephone:='';
      if aSuperArray[i].AsObject.Exists('�ⲿϵͳΨһ���') then His_Unid:=aSuperArray[i]['�ⲿϵͳΨһ���'].AsString else His_Unid:='';
      if aSuperArray[i].AsObject.Exists('���뵥���') then DNH:=aSuperArray[i]['���뵥���'].AsString else DNH:='';

      if 'Excel'=aJson['JSON����Դ'].AsString then chk_con_unid:=ScalarSQLCmd(AAdoconnstr,'select top 1 unid from chk_con where patientname='''+patientname+''' AND sex='''+sex+''' AND age='''+age+''' AND combin_id='''+WorkGroup+''' and isnull(report_doctor,'''')='''' ')
        else chk_con_unid:=ScalarSQLCmd(AAdoconnstr,'select top 1 unid from chk_con cc where cc.combin_id='''+WorkGroup+''' and cc.TjJianYan='''+aSuperArrayMX[j]['�����'].AsString+''' and cc.flagetype='''+SampleType+''' and isnull(report_doctor,'''')='''' ');
        
      if chk_con_unid='' then//���ڹ����顢����š�����������ͬ,��δ��˵ļ��鵥,���ڴ˼��鵥��������ϸ.���������һ�����鵥
      begin
        lsh:=ScalarSQLCmd(AAdoconnstr,' select dbo.uf_GetNextSerialNum('''+WorkGroup+''','''+FormatDateTime('YYYY-MM-DD',ServerDateTime)+''','''+YXJB+''') ');

        adoconn11:=Tadoconnection.Create(nil);
        adoconn11.ConnectionString:=AAdoconnstr;
        adoconn11.LoginPrompt:=false;

        adotemp11:=Tadoquery.Create(nil);
        adotemp11.Connection:=adoconn11;
        adotemp11.Close;
        adotemp11.SQL.Clear;
        adotemp11.SQL.Add('insert into chk_con ( combin_id, checkid, patientname, sex, age, Caseno, report_date, deptname, check_doctor, His_Unid, Diagnosetype, flagetype, typeflagcase, LSH,');
        adotemp11.SQL.Add(' bedno, diagnose, issure, WorkCompany, WorkDepartment, WorkCategory, WorkID, ifMarry, OldAddress, Address, Telephone, DNH, TjJianYan) values ');
        adotemp11.SQL.Add('                    (:combin_id,:checkid,:patientname,:sex,:age,:Caseno,:report_date,:deptname,:check_doctor,:His_Unid,:Diagnosetype,:flagetype,:typeflagcase,:LSH,');
        adotemp11.SQL.Add(':bedno,:diagnose,:issure,:WorkCompany,:WorkDepartment,:WorkCategory,:WorkID,:ifMarry,:OldAddress,:Address,:Telephone,:DNH,:TjJianYan)');
        adotemp11.SQL.Add(' SELECT SCOPE_IDENTITY() AS Insert_Identity ');
        adotemp11.Parameters.ParamByName('combin_id').Value:=WorkGroup;
        adotemp11.Parameters.ParamByName('checkid').Value:=checkid;
        adotemp11.Parameters.ParamByName('patientname').Value:=patientname;
        adotemp11.Parameters.ParamByName('sex').Value:=sex;
        adotemp11.Parameters.ParamByName('age').Value:=age;
        adotemp11.Parameters.ParamByName('Caseno').Value:=Caseno;
        adotemp11.Parameters.ParamByName('report_date').Value:=RequestDate;
        adotemp11.Parameters.ParamByName('deptname').Value:=deptname;
        adotemp11.Parameters.ParamByName('check_doctor').Value:=check_doctor;
        adotemp11.Parameters.ParamByName('His_Unid').Value:=His_Unid;
        adotemp11.Parameters.ParamByName('Diagnosetype').Value:=YXJB;
        adotemp11.Parameters.ParamByName('flagetype').Value:=SampleType;
        adotemp11.Parameters.ParamByName('typeflagcase').Value:=SampleStatus;
        adotemp11.Parameters.ParamByName('LSH').Value:=lsh;
        adotemp11.Parameters.ParamByName('bedno').Value:=bedno;
        adotemp11.Parameters.ParamByName('diagnose').Value:=diagnose;
        adotemp11.Parameters.ParamByName('issure').Value:=issure;
        adotemp11.Parameters.ParamByName('WorkCompany').Value:=WorkCompany;
        adotemp11.Parameters.ParamByName('WorkDepartment').Value:=WorkDepartment;
        adotemp11.Parameters.ParamByName('WorkCategory').Value:=WorkCategory;
        adotemp11.Parameters.ParamByName('WorkID').Value:=WorkID;
        adotemp11.Parameters.ParamByName('ifMarry').Value:=ifMarry;
        adotemp11.Parameters.ParamByName('OldAddress').Value:=OldAddress;
        adotemp11.Parameters.ParamByName('Address').Value:=Address;
        adotemp11.Parameters.ParamByName('Telephone').Value:=Telephone;
        adotemp11.Parameters.ParamByName('DNH').Value:=DNH;
        if 'Excel'=aJson['JSON����Դ'].AsString then
          adotemp11.Parameters.ParamByName('TjJianYan').Value:=''
        else
          adotemp11.Parameters.ParamByName('TjJianYan').Value:=aSuperArrayMX[j]['�����'].AsString;
        Try
          adotemp11.Open;
        except
          on E:Exception do
          begin
            MessageDlg('���벡����Ϣʧ��:'+E.Message,mtError,[MBOK],0);
            adotemp11.Free;
            adoconn11.Free;
            exit;
          end;
        end;
        chk_con_unid:=adotemp11.fieldbyname('Insert_Identity').AsString;
        adotemp11.Free;
        adoconn11.Free;
      end;

      //������ϸbegin
      ExecSQLCmd(AAdoconnstr,'update chk_valu set issure=1 where pkunid='+chk_con_unid+' and pkcombin_id='''+pkcombin_id+''' and isnull(issure,'''')<>''1'' ');

      adoconn22:=Tadoconnection.Create(nil);
      adoconn22.ConnectionString:=strpas(AAdoconnstr);
      adoconn22.LoginPrompt:=false;

      adotemp22:=Tadoquery.Create(nil);
      adotemp22.Connection:=adoconn22;
      adotemp22.Close;
      adotemp22.SQL.Clear;
      adotemp22.SQL.Text:='select cci.itemid from CombSChkItem csci,combinitem ci,clinicchkitem cci '+
        ' where csci.CombUnid=ci.Unid and cci.unid=csci.ItemUnid and ci.Id='''+pkcombin_id+''' ';
      Try
        adotemp22.Open;
      except
        on E:Exception do
        begin
          MessageDlg('��ȡָ�������Ŀ������Ŀʧ��:'+E.Message,mtError,[MBOK],0);
          adotemp22.Free;
          adoconn22.Free;
          exit;
        end;
      end;
      while not adotemp22.Eof do
      begin
        if '1'<>ScalarSQLCmd(AAdoconnstr,'select top 1 1 from chk_valu where pkunid='+chk_con_unid+' and pkcombin_id='''+aSuperArrayMX[j]['LIS�����Ŀ����'].AsString+''' and itemid='''+adotemp22.FieldByName('itemid').AsString+''' ') then
          ExecSQLCmd(AAdoconnstr,'insert into chk_valu (pkunid,pkcombin_id,itemid,issure) values ('+chk_con_unid+','''+aSuperArrayMX[j]['LIS�����Ŀ����'].AsString+''','''+adotemp22.FieldByName('itemid').AsString+''',1)');

        adotemp22.Next;
      end;
      adotemp22.Free;
      adoconn22.Free;

      //Data2Lis������ʱҲ����ã��ʴ˴���ע��
      //addOrEditCalcItem(pchar(LisConn),pchar(s2),checkunid);//���Ӽ�����Ŀ
      //addOrEditCalcValu(pchar(LisConn),checkunid,false,'');//���¼�����Ŀ
      //������ϸend
    end;
  end;
end;

exports
  RequestForm2Lis;

begin
end.
