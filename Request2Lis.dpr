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
//��DLL�ṩ����RequestForm2Lis�����ڽ�JSON��ʽ�ļ���������Ϣ����LIS
//3���������:
//AAdoconnstr:LIS���ݿ�������ַ���
//ARequestJSON:JSON��ʽ�ļ������뵥��Ϣ
//CurrentWorkGroup:��ǰ������.�������Ŀδ����Ĭ�Ϲ�����,���뵥�����뵱ǰ������
//JSONʵ��:
//  {
//      "JSON����Դ":"HIS",
//      "����ҽ��": [
//          {
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
//                      "���뵥���": "10000",
//                      "���ȼ���": "����",
//                      "��������": "Ѫ��",
//                      "����״̬": "����"
//                  },
//                  {
//                      "������": "X0013",
//                      "LIS�����Ŀ����": "54",
//                      "�����": "12346",
//                      "���뵥���": "10001",
//                      "���ȼ���": "����",
//                      "��������": "ȫѪ",
//                      "����״̬": "����"
//                  }
//              ]
//          },
//          {
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
//                      "���뵥���": "10002",
//                      "���ȼ���": "����",
//                      "��������": "Ѫ��",
//                      "����״̬": "����"
//                  },
//                  {
//                      "������": "X0014",
//                      "LIS�����Ŀ����": "54",
//                      "�����": "12348",
//                      "���뵥���": "10003",
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
//�����뵥��š�:ÿ��һ�������Ŀ,���ɵ�һ��Ψһ����,��ЩHIS(������PEIS)�ô˺���ƥ�������Ŀ�µ�����Ŀ
//
//2023-02-17�������Ѹ��ݹ����顢��������Ϊ���ݽ��в�
//�Ƿ�Ҫ��������Ŀ��������ĸ�����в𵥣��۲�Ӧ������ٶ�
//
//��JSON��ʾΪ��ͼ����վ:https://jsoncrack.com/editor
//==============================================================================
//��DLL�ṩ����GetLisCombItem�����ڷ���JSON��ʽ��LIS�����Ŀ��Ϣ
//�������:
//AAdoconnstr:LIS���ݿ�������ַ���
//AHisItem:HIS�����Ŀ�б�,�ö��ŷָ�
//AEquipWord:������ĸ
//AExtSystemId:�ⲿϵͳID
//����JSON��ʽ����
//{
//  "��Ŀ��Ϣ": [
//    {
//      "�����ĿUNID": 123,
//      "�����Ŀ����": "01",
//      "�����Ŀ����": "",
//      "�����Ŀ��ע": "",
//      "�����ĿĬ�Ϲ�����": "",
//      "�����ĿĬ����������": "",
//      "�����Ŀ�����ָ���": ""
//    },
//    {
//      "�����ĿUNID": 124,
//      "�����Ŀ����": "02",
//      "�����Ŀ����": "",
//      "�����Ŀ��ע": "",
//      "�����ĿĬ�Ϲ�����": "",
//      "�����ĿĬ����������": "",
//      "�����Ŀ�����ָ���": ""
//    }
//  ]
//}
//==============================================================================
//��DLL�ṩ����GetLisSubItem�����ڷ���JSON��ʽ��LIS����Ŀ��Ϣ
//�������:
//AAdoconnstr:LIS���ݿ�������ַ���
//AHisItem:HIS�����Ŀ�б�,�ö��ŷָ�
//AEquipWord:������ĸ
//AExtSystemId:�ⲿϵͳID
//����JSON��ʽ����
//{
//  "��Ŀ��Ϣ": [
//        {
//          "����ĿUNID": 125,
//          "����Ŀ����": "1011",
//          "����Ŀ����": "",
//          "����ĿӢ����": "",
//          "����Ŀ������ʶ": "",
//          "����Ŀ�����ֶ�1": "",
//          "����Ŀ�����ֶ�2": "",
//          "����Ŀ�����ֶ�3": "",
//          "����Ŀ�����ֶ�4": "",
//          "����Ŀ�����ֶ�5": "",
//          "����Ŀ�����ֶ�6": "",
//          "����Ŀ�����ֶ�7": "",
//          "����Ŀ�����ֶ�8": "",
//          "����Ŀ�����ֶ�9": "",
//          "����Ŀ�����ֶ�10": "",
//          "����Ŀ����������ʶ": ""
//        },
//        {
//          "����ĿUNID": 126,
//          "����Ŀ����": "1012",
//          "����Ŀ����": "",
//          "����ĿӢ����": "",
//          "����Ŀ������ʶ": "",
//          "����Ŀ�����ֶ�1": "",
//          "����Ŀ�����ֶ�2": "",
//          "����Ŀ�����ֶ�3": "",
//          "����Ŀ�����ֶ�4": "",
//          "����Ŀ�����ֶ�5": "",
//          "����Ŀ�����ֶ�6": "",
//          "����Ŀ�����ֶ�7": "",
//          "����Ŀ�����ֶ�8": "",
//          "����Ŀ�����ֶ�9": "",
//          "����Ŀ�����ֶ�10": "",
//          "����Ŀ����������ʶ": ""
//        }
//  ]
//}
//==============================================================================

uses
  SysUtils,
  Classes,
  ADODB,
  DateUtils,
  Dialogs,
  superobject in 'superobject.pas';

{$R *.res}
function UnicodeToChinese(const AUnicodeStr:PChar):PChar;stdcall;external 'LYFunction.dll';

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
  Surem1:String;//���뵥���(HIS)
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

      if aSuperArrayMX[j].AsObject.Exists('LIS�����Ŀ����') then pkcombin_id:=aSuperArrayMX[j].S['LIS�����Ŀ����'] else pkcombin_id:=''; 
      if pkcombin_id='' then pkcombin_id:='�����ڵ������Ŀ����';
        
      defaultWorkGroup:=ScalarSQLCmd(AAdoconnstr,'select dept_DfValue from combinitem where Id='''+pkcombin_id+''' ');
      defaultSampleType:=ScalarSQLCmd(AAdoconnstr,'select specimentype_DfValue from combinitem where Id='''+pkcombin_id+''' ');

      //���Ĭ�Ϲ�����Ϊ��,���뵱ǰ������
      WorkGroup:=defaultWorkGroup;
      if WorkGroup='' then WorkGroup:=CurrentWorkGroup;
      if 'Excel'=aJson.S['JSON����Դ'] then WorkGroup:=CurrentWorkGroup;

      //���JSON����������Ϊ��,��ȡĬ����������
      if aSuperArrayMX[j].AsObject.Exists('��������') then SampleType:=aSuperArrayMX[j].S['��������'] else SampleType:=''; 
      if (SampleType='')and('Excel'<>aJson.S['JSON����Դ']) then SampleType:=defaultSampleType;

      if aSuperArrayMX[j].AsObject.Exists('���ȼ���') then YXJB:=aSuperArrayMX[j].S['���ȼ���'] else YXJB:=''; 
      if YXJB='' then YXJB:='����';

      if aSuperArrayMX[j].AsObject.Exists('����״̬') then SampleStatus:=aSuperArrayMX[j].S['����״̬'] else SampleStatus:=''; 
      if SampleStatus='' then SampleStatus:='����';

      fs.DateSeparator:='-';
      fs.TimeSeparator:=':';
      fs.ShortDateFormat:='YYYY-MM-DD hh:nn:ss';
      if aSuperArray[i].AsObject.Exists('��������') then RequestDateStr:=aSuperArray[i].S['��������'] else RequestDateStr:='';
      RequestDate:=StrtoDateTimeDef(RequestDateStr,ServerDateTime,fs);
      if  RequestDate<2 then ReplaceDate(RequestDate,ServerDateTime);//��ʾ1899-12-30,û�и����ڸ�ֵ
      if (HourOf(RequestDate)=0) and (MinuteOf(RequestDate)=0) and (SecondOf(RequestDate)=0) then ReplaceTime(RequestDate,ServerDateTime);//��ʾû�и�ʱ�丳ֵ

      if aSuperArray[i].AsObject.Exists('��������') then patientname:=aSuperArray[i].S['��������'] else patientname:='';
      if aSuperArray[i].AsObject.Exists('�����Ա�') then sex:=aSuperArray[i].S['�����Ա�'] else sex:=''; 
      if aSuperArray[i].AsObject.Exists('��������') then age:=aSuperArray[i].S['��������'] else age:='';
      if aSuperArray[i].AsObject.Exists('������') then Caseno:=aSuperArray[i].S['������'] else Caseno:='';
      if aSuperArray[i].AsObject.Exists('�������') then deptname:=aSuperArray[i].S['�������'] else deptname:='';
      if aSuperArray[i].AsObject.Exists('����ҽ��') then check_doctor:=aSuperArray[i].S['����ҽ��'] else check_doctor:='';
      if aSuperArray[i].AsObject.Exists('����') then bedno:=aSuperArray[i].S['����'] else bedno:='';
      if aSuperArray[i].AsObject.Exists('�ٴ����') then diagnose:=aSuperArray[i].S['�ٴ����'] else diagnose:='';
      if aSuperArray[i].AsObject.Exists('��ע') then issure:=aSuperArray[i].S['��ע'] else issure:='';
      if aSuperArray[i].AsObject.Exists('������˾') then WorkCompany:=aSuperArray[i].S['������˾'] else WorkCompany:='';
      if aSuperArray[i].AsObject.Exists('��������') then WorkDepartment:=aSuperArray[i].S['��������'] else WorkDepartment:='';
      if aSuperArray[i].AsObject.Exists('����') then WorkCategory:=aSuperArray[i].S['����'] else WorkCategory:='';
      if aSuperArray[i].AsObject.Exists('����') then WorkID:=aSuperArray[i].S['����'] else WorkID:='';
      if aSuperArray[i].AsObject.Exists('���') then ifMarry:=aSuperArray[i].S['���'] else ifMarry:='';
      if aSuperArray[i].AsObject.Exists('����') then OldAddress:=aSuperArray[i].S['����'] else OldAddress:='';
      if aSuperArray[i].AsObject.Exists('סַ') then Address:=aSuperArray[i].S['סַ'] else Address:='';
      if aSuperArray[i].AsObject.Exists('�绰') then Telephone:=aSuperArray[i].S['�绰'] else Telephone:='';
      if aSuperArray[i].AsObject.Exists('�ⲿϵͳΨһ���') then His_Unid:=aSuperArray[i].S['�ⲿϵͳΨһ���'] else His_Unid:='';
      if aSuperArrayMX[j].AsObject.Exists('���뵥���') then Surem1:=aSuperArrayMX[j].S['���뵥���'] else Surem1:='';
      if aSuperArrayMX[j].AsObject.Exists('������') then checkid:=aSuperArrayMX[j].S['������'] else checkid:='';

      if 'Excel'=aJson.S['JSON����Դ'] then chk_con_unid:=ScalarSQLCmd(AAdoconnstr,'select top 1 unid from chk_con where patientname='''+patientname+''' AND sex='''+sex+''' AND age='''+age+''' AND combin_id='''+WorkGroup+''' and isnull(report_doctor,'''')='''' ')
        else chk_con_unid:=ScalarSQLCmd(AAdoconnstr,'select top 1 unid from chk_con cc where cc.combin_id='''+WorkGroup+''' and cc.TjJianYan='''+aSuperArrayMX[j].S['�����']+''' and cc.flagetype='''+SampleType+''' and isnull(report_doctor,'''')='''' ');
        
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
        adotemp11.SQL.Add(' bedno, diagnose, issure, WorkCompany, WorkDepartment, WorkCategory, WorkID, ifMarry, OldAddress, Address, Telephone, TjJianYan) values ');
        adotemp11.SQL.Add('                    (:combin_id,:checkid,:patientname,:sex,:age,:Caseno,:report_date,:deptname,:check_doctor,:His_Unid,:Diagnosetype,:flagetype,:typeflagcase,:LSH,');
        adotemp11.SQL.Add(':bedno,:diagnose,:issure,:WorkCompany,:WorkDepartment,:WorkCategory,:WorkID,:ifMarry,:OldAddress,:Address,:Telephone,:TjJianYan)');
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
        //adotemp11.Parameters.ParamByName('DNH').Value:=DNH;
        if 'Excel'=aJson['JSON����Դ'].AsString then
          adotemp11.Parameters.ParamByName('TjJianYan').Value:=''
        else
          adotemp11.Parameters.ParamByName('TjJianYan').Value:=aSuperArrayMX[j].S['�����'];
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
        if '1'<>ScalarSQLCmd(AAdoconnstr,'select top 1 1 from chk_valu where pkunid='+chk_con_unid+' and pkcombin_id='''+aSuperArrayMX[j].S['LIS�����Ŀ����']+''' and itemid='''+adotemp22.FieldByName('itemid').AsString+''' ') then
          ExecSQLCmd(AAdoconnstr,'insert into chk_valu (pkunid,pkcombin_id,itemid,Surem1,issure) values ('+chk_con_unid+','''+aSuperArrayMX[j].S['LIS�����Ŀ����']+''','''+adotemp22.FieldByName('itemid').AsString+''','''+Surem1+''',1)');

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

function GetLisCombItem(const AAdoconnstr,AHisItem,AEquipWord,AExtSystemId:PChar):PChar;stdcall;
var
  adoconn:Tadoconnection;
  adotemp22:Tadoquery;
  ls:TStrings;
  i,j:Integer;
  
  ObjectCombItem:ISuperObject;
  ArrayCombItem:ISuperObject;

  ResultObject:ISuperObject;

  ifExistsKeyValue:boolean;
begin
  ResultObject:=SO;

  ls:=TStringList.Create;
  ExtractStrings([','],[],AHisItem,ls);

  for i := 0 to ls.Count-1 do
  begin
    if trim(ls[i])='' then continue;
    
    adoconn:=Tadoconnection.Create(nil);
    adoconn.ConnectionString:=Aadoconnstr;
    adoconn.LoginPrompt:=false;

    adotemp22:=Tadoquery.Create(nil);
    adotemp22.Connection:=adoconn;
    adotemp22.Close;
    adotemp22.SQL.Clear;
    adotemp22.SQL.Text:='select ci.Unid,ci.Id,ci.Name,ci.Remark,ci.dept_DfValue,ci.specimentype_DfValue,ci.itemtype '+
                        'from combinitem ci,HisCombItem hci,CombSChkItem csci,clinicchkitem cci '+
                        'where ci.Unid=hci.CombUnid and hci.ExtSystemId='''+AExtSystemId+
                        ''' and hci.HisItem='''+ls[i]+
                        ''' and csci.CombUnid=ci.Unid '+
                        'and cci.unid=csci.ItemUnid '+ 
                        'and cci.COMMWORD='''+AEquipWord+
                        ''' group by ci.Unid,ci.Id,ci.Name,ci.Remark,ci.dept_DfValue,ci.specimentype_DfValue,ci.itemtype';
    Try
      adotemp22.Open;
    except
      on E:Exception do
      begin
        adotemp22.Free;
        adoconn.Free;
        
        continue;
      end;
    end;

    while not adotemp22.Eof do
    begin
      if not ResultObject.AsObject.Exists('��Ŀ��Ϣ') then
      begin
        ObjectCombItem:=SO;
        ObjectCombItem.I['�����ĿUNID'] := adotemp22.fieldbyname('Unid').AsInteger;
        ObjectCombItem.S['�����Ŀ����'] := adotemp22.fieldbyname('Id').AsString;
        ObjectCombItem.S['�����Ŀ����'] := adotemp22.fieldbyname('Name').AsString;
        ObjectCombItem.S['�����Ŀ��ע'] := adotemp22.fieldbyname('Remark').AsString;
        ObjectCombItem.S['�����ĿĬ�Ϲ�����'] := adotemp22.fieldbyname('dept_DfValue').AsString;
        ObjectCombItem.S['�����ĿĬ����������'] := adotemp22.fieldbyname('specimentype_DfValue').AsString;
        ObjectCombItem.S['�����Ŀ�����ָ���'] := adotemp22.fieldbyname('itemtype').AsString;

        ArrayCombItem:=SA([]);
        ArrayCombItem.AsArray.Add(ObjectCombItem);
        ObjectCombItem:=nil;

        ResultObject.O['��Ŀ��Ϣ']:=ArrayCombItem;
        ArrayCombItem:=nil;

        adotemp22.Next;
        continue;
      end;

      ifExistsKeyValue:=false;
      for j:=0 to ResultObject['��Ŀ��Ϣ'].AsArray.Length-1 do
      begin
        if ResultObject['��Ŀ��Ϣ'].AsArray[j].I['�����ĿUNID']=adotemp22.fieldbyname('Unid').AsInteger then
        begin
          ifExistsKeyValue:=true;
          break;
        end;
      end;

      if ifExistsKeyValue then begin adotemp22.Next;continue;end;

      ObjectCombItem:=SO;
      ObjectCombItem.I['�����ĿUNID'] := adotemp22.fieldbyname('Unid').AsInteger;
      ObjectCombItem.S['�����Ŀ����'] := adotemp22.fieldbyname('Id').AsString;
      ObjectCombItem.S['�����Ŀ����'] := adotemp22.fieldbyname('Name').AsString;
      ObjectCombItem.S['�����Ŀ��ע'] := adotemp22.fieldbyname('Remark').AsString;
      ObjectCombItem.S['�����ĿĬ�Ϲ�����'] := adotemp22.fieldbyname('dept_DfValue').AsString;
      ObjectCombItem.S['�����ĿĬ����������'] := adotemp22.fieldbyname('specimentype_DfValue').AsString;
      ObjectCombItem.S['�����Ŀ�����ָ���'] := adotemp22.fieldbyname('itemtype').AsString;

      ResultObject.O['��Ŀ��Ϣ'].AsArray.Add(ObjectCombItem);
      ObjectCombItem:=nil;

      adotemp22.Next;
    end;
    
    adotemp22.Free;
    adoconn.Free;
  end;

  ls.Free;

  Result:=UnicodeToChinese(PChar(AnsiString(ResultObject.AsJson)));

  ResultObject:=nil;
end;

function GetLisSubItem(const AAdoconnstr,AHisItem,AEquipWord,AExtSystemId:PChar):PChar;stdcall;
var
  adoconn:Tadoconnection;
  adotemp22:Tadoquery;
  ls:TStrings;
  i,j:Integer;
  
  ObjectSubItem:ISuperObject;
  ArraySubItem:ISuperObject;

  ResultObject:ISuperObject;

  ifExistsKeyValue:boolean;
begin
  ResultObject:=SO;

  ls:=TStringList.Create;
  ExtractStrings([','],[],AHisItem,ls);

  for i := 0 to ls.Count-1 do
  begin
    if trim(ls[i])='' then continue;
    
    adoconn:=Tadoconnection.Create(nil);
    adoconn.ConnectionString:=Aadoconnstr;
    adoconn.LoginPrompt:=false;

    adotemp22:=Tadoquery.Create(nil);
    adotemp22.Connection:=adoconn;
    adotemp22.Close;
    adotemp22.SQL.Clear;
    adotemp22.SQL.Text:='select cci.unid,cci.itemid,cci.name,cci.english_name,cci.dlttype,cci.Reserve1,cci.Reserve2,cci.Dosage1,cci.Dosage2,cci.Reserve5,cci.Reserve6,cci.Reserve7,cci.Reserve8,cci.Reserve9,cci.Reserve10,cci.defaultvalue '+ 
                        'from combinitem ci,HisCombItem hci,CombSChkItem csci,clinicchkitem cci '+
                        'where ci.Unid=hci.CombUnid and hci.ExtSystemId='''+AExtSystemId+
                        ''' and hci.HisItem='''+ls[i]+
                        ''' and csci.CombUnid=ci.Unid '+
                        'and cci.unid=csci.ItemUnid '+ 
                        'and cci.COMMWORD='''+AEquipWord+
                        ''' group by cci.unid,cci.itemid,cci.name,cci.english_name,cci.dlttype,cci.Reserve1,cci.Reserve2,cci.Dosage1,cci.Dosage2,cci.Reserve5,cci.Reserve6,cci.Reserve7,cci.Reserve8,cci.Reserve9,cci.Reserve10,cci.defaultvalue';
    Try
      adotemp22.Open;
    except
      on E:Exception do
      begin
        adotemp22.Free;
        adoconn.Free;
        
        continue;
      end;
    end;

    while not adotemp22.Eof do
    begin
      if not ResultObject.AsObject.Exists('��Ŀ��Ϣ') then
      begin
        ObjectSubItem:=SO;
        ObjectSubItem.I['����ĿUNID'] := adotemp22.fieldbyname('unid').AsInteger;
        ObjectSubItem.S['����Ŀ����'] := adotemp22.fieldbyname('itemid').AsString;
        ObjectSubItem.S['����Ŀ����'] := adotemp22.fieldbyname('name').AsString;
        ObjectSubItem.S['����ĿӢ����'] := adotemp22.fieldbyname('english_name').AsString;
        ObjectSubItem.S['����Ŀ������ʶ'] := adotemp22.fieldbyname('dlttype').AsString;
        ObjectSubItem.S['����Ŀ�����ֶ�1'] := adotemp22.fieldbyname('Reserve1').AsString;
        ObjectSubItem.S['����Ŀ�����ֶ�2'] := adotemp22.fieldbyname('Reserve2').AsString;
        ObjectSubItem.S['����Ŀ�����ֶ�3'] := adotemp22.fieldbyname('Dosage1').AsString;
        ObjectSubItem.S['����Ŀ�����ֶ�4'] := adotemp22.fieldbyname('Dosage2').AsString;
        ObjectSubItem.S['����Ŀ�����ֶ�5'] := adotemp22.fieldbyname('Reserve5').AsString;
        ObjectSubItem.S['����Ŀ�����ֶ�6'] := adotemp22.fieldbyname('Reserve6').AsString;
        ObjectSubItem.S['����Ŀ�����ֶ�7'] := adotemp22.fieldbyname('Reserve7').AsString;
        ObjectSubItem.S['����Ŀ�����ֶ�8'] := adotemp22.fieldbyname('Reserve8').AsString;
        ObjectSubItem.S['����Ŀ�����ֶ�9'] := adotemp22.fieldbyname('Reserve9').AsString;
        ObjectSubItem.S['����Ŀ�����ֶ�10'] := adotemp22.fieldbyname('Reserve10').AsString;
        ObjectSubItem.S['����Ŀ����������ʶ'] := adotemp22.fieldbyname('defaultvalue').AsString;

        ArraySubItem:=SA([]);
        ArraySubItem.AsArray.Add(ArraySubItem);
        ObjectSubItem:=nil;

        ResultObject.O['��Ŀ��Ϣ']:=ArraySubItem;
        ArraySubItem:=nil;

        adotemp22.Next;
        continue;
      end;

      ifExistsKeyValue:=false;
      for j:=0 to ResultObject['��Ŀ��Ϣ'].AsArray.Length-1 do
      begin
        if ResultObject['��Ŀ��Ϣ'].AsArray[j].I['�����ĿUNID']=adotemp22.fieldbyname('Unid').AsInteger then
        begin
          ifExistsKeyValue:=true;
          break;
        end;
      end;

      if ifExistsKeyValue then begin adotemp22.Next;continue;end;

      ObjectSubItem:=SO;
      ObjectSubItem.I['����ĿUNID'] := adotemp22.fieldbyname('unid').AsInteger;
      ObjectSubItem.S['����Ŀ����'] := adotemp22.fieldbyname('itemid').AsString;
      ObjectSubItem.S['����Ŀ����'] := adotemp22.fieldbyname('name').AsString;
      ObjectSubItem.S['����ĿӢ����'] := adotemp22.fieldbyname('english_name').AsString;
      ObjectSubItem.S['����Ŀ������ʶ'] := adotemp22.fieldbyname('dlttype').AsString;
      ObjectSubItem.S['����Ŀ�����ֶ�1'] := adotemp22.fieldbyname('Reserve1').AsString;
      ObjectSubItem.S['����Ŀ�����ֶ�2'] := adotemp22.fieldbyname('Reserve2').AsString;
      ObjectSubItem.S['����Ŀ�����ֶ�3'] := adotemp22.fieldbyname('Dosage1').AsString;
      ObjectSubItem.S['����Ŀ�����ֶ�4'] := adotemp22.fieldbyname('Dosage2').AsString;
      ObjectSubItem.S['����Ŀ�����ֶ�5'] := adotemp22.fieldbyname('Reserve5').AsString;
      ObjectSubItem.S['����Ŀ�����ֶ�6'] := adotemp22.fieldbyname('Reserve6').AsString;
      ObjectSubItem.S['����Ŀ�����ֶ�7'] := adotemp22.fieldbyname('Reserve7').AsString;
      ObjectSubItem.S['����Ŀ�����ֶ�8'] := adotemp22.fieldbyname('Reserve8').AsString;
      ObjectSubItem.S['����Ŀ�����ֶ�9'] := adotemp22.fieldbyname('Reserve9').AsString;
      ObjectSubItem.S['����Ŀ�����ֶ�10'] := adotemp22.fieldbyname('Reserve10').AsString;
      ObjectSubItem.S['����Ŀ����������ʶ'] := adotemp22.fieldbyname('defaultvalue').AsString;

      ResultObject.O['��Ŀ��Ϣ'].AsArray.Add(ObjectSubItem);
      ObjectSubItem:=nil;

      adotemp22.Next;
    end;
    
    adotemp22.Free;
    adoconn.Free;
  end;

  ls.Free;

  Result:=UnicodeToChinese(PChar(AnsiString(ResultObject.AsJson)));

  ResultObject:=nil;
end;

exports
  RequestForm2Lis,
  GetLisCombItem,
  GetLisSubItem;

begin
end.
