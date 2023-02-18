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
//      "����ҽ��": [
//          {
//              "ҽ��Ψһ���": "10000",
//              "������": "101234",
//              "��������": "�ܲ�",
//              "�����Ա�": "��",
//              "��������": "24��",
//              "��������": "2023-2-6",
//              "�������": "�ڿ�",
//              "����ҽ��": "���",
//              "ҽ����ϸ": [
//                  {
//                      "������": "S0087",
//                      "LIS���������Ŀ����": "06",
//                      "���ȼ���": "����",
//                      "��������": "Ѫ��",
//                      "����״̬": "����"
//                  },
//                  {
//                      "������": "X0013",
//                      "LIS���������Ŀ����": "54",
//                      "���ȼ���": "����",
//                      "��������": "ȫѪ",
//                      "����״̬": "����"
//                  }
//              ]
//          },
//          {
//              "ҽ��Ψһ���": "10001",
//              "������": "101221",
//              "��������": "����",
//              "�����Ա�": "��",
//              "��������": "25��",
//              "��������": "2023-2-7",
//              "�������": "���",
//              "����ҽ��": "�ŷ�",
//              "ҽ����ϸ": [
//                  {
//                      "������": "S0088",
//                      "LIS���������Ŀ����": "06",
//                      "���ȼ���": "����",
//                      "��������": "Ѫ��",
//                      "����״̬": "����"
//                  },
//                  {
//                      "������": "X0014",
//                      "LIS���������Ŀ����": "54",
//                      "���ȼ���": "����",
//                      "��������": "ȫѪ",
//                      "����״̬": "����"
//                  }
//              ]
//          }
//      ]
//  }
//����JSON�����ֶα������
//ֵ������ֶΣ�ҽ��Ψһ��š���ҽ��Ψһ��š�����HIS���ؼ������ı�ʶ,���ǳ���������Ŀ����ͬһ�ż��鵥���ж�����
//JSON������ʱ���ʽ��YYYY-MM-DD hh:nn:ss
//�����LIS���������Ŀ���롿��ֵ��LIS�в����ڣ���ֻ�ᵼ�벡�˻�����Ϣ�����ᵼ�������Ŀ
//
//2023-02-17�������Ѹ��ݹ����顢��������Ϊ���ݽ��в�
//�Ƿ�Ҫ��������Ŀ��������ĸ�����в𵥣��۲�Ӧ������ٶ�
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
  lsh:string;
begin
  ServerDateTime:=GetServerDate(AAdoconnstr);

  aJson:=SO(ARequestJSON);
  aSuperArray:=aJson['����ҽ��'].AsArray;
  for i:=0 to aSuperArray.Length-1 do
  begin
    aSuperArrayMX:=aSuperArray[i]['ҽ����ϸ'].AsArray;
    for j:=0 to aSuperArrayMX.Length-1 do
    begin
      defaultWorkGroup:=ScalarSQLCmd(AAdoconnstr,'select dept_DfValue from combinitem where Id='''+aSuperArrayMX[j]['LIS���������Ŀ����'].AsString+''' ');
      defaultSampleType:=ScalarSQLCmd(AAdoconnstr,'select specimentype_DfValue from combinitem where Id='''+aSuperArrayMX[j]['LIS���������Ŀ����'].AsString+''' ');

      //���Ĭ�Ϲ�����Ϊ��,���뵱ǰ������
      WorkGroup:=defaultWorkGroup;
      if WorkGroup='' then WorkGroup:=CurrentWorkGroup;

      //���JSON����������Ϊ��,��ȡĬ����������
      SampleType:=aSuperArrayMX[j]['��������'].AsString;
      if SampleType='' then SampleType:=defaultSampleType;

      YXJB:=aSuperArrayMX[j]['���ȼ���'].AsString;
      if YXJB='' then YXJB:='����';

      SampleStatus:=aSuperArrayMX[j]['����״̬'].AsString;
      if SampleStatus='' then SampleStatus:='����';

      fs.DateSeparator:='-';
      fs.TimeSeparator:=':';
      fs.ShortDateFormat:='YYYY-MM-DD hh:nn:ss';
      RequestDate:=StrtoDateTimeDef(aSuperArray[i]['��������'].AsString,ServerDateTime,fs);
      if  RequestDate<2 then ReplaceDate(RequestDate,ServerDateTime);//��ʾ1899-12-30,û�и����ڸ�ֵ
      if (HourOf(RequestDate)=0) and (MinuteOf(RequestDate)=0) and (SecondOf(RequestDate)=0) then ReplaceTime(RequestDate,ServerDateTime);//��ʾû�и�ʱ�丳ֵ

      chk_con_unid:=ScalarSQLCmd(AAdoconnstr,'select top 1 unid from chk_con cc where cc.combin_id='''+WorkGroup+''' and cc.His_Unid='''+aSuperArray[i]['ҽ��Ψһ���'].AsString+''' and cc.flagetype='''+SampleType+''' and isnull(report_doctor,'''')='''' ');
      if chk_con_unid='' then//���ڹ����顢ҽ��Ψһ��š�����������ͬ,��δ��˵ļ��鵥,���ڴ˼��鵥��������ϸ.���������һ�����鵥
      begin
        lsh:=ScalarSQLCmd(AAdoconnstr,' select dbo.uf_GetNextSerialNum('''+WorkGroup+''','''+FormatDateTime('YYYY-MM-DD',ServerDateTime)+''','''+YXJB+''') ');

        adoconn11:=Tadoconnection.Create(nil);
        adoconn11.ConnectionString:=AAdoconnstr;
        adoconn11.LoginPrompt:=false;

        adotemp11:=Tadoquery.Create(nil);
        adotemp11.Connection:=adoconn11;
        adotemp11.Close;
        adotemp11.SQL.Clear;
        adotemp11.SQL.Add('insert into chk_con ( combin_id, checkid, patientname, sex, age, Caseno, report_date, deptname, check_doctor, His_Unid, Diagnosetype, flagetype, typeflagcase, LSH) values ');
        adotemp11.SQL.Add('                    (:combin_id,:checkid,:patientname,:sex,:age,:Caseno,:report_date,:deptname,:check_doctor,:His_Unid,:Diagnosetype,:flagetype,:typeflagcase,:LSH)');
        adotemp11.SQL.Add(' SELECT SCOPE_IDENTITY() AS Insert_Identity ');
        adotemp11.Parameters.ParamByName('combin_id').Value:=WorkGroup;
        adotemp11.Parameters.ParamByName('checkid').Value:=aSuperArrayMX[j]['������'].AsString;
        adotemp11.Parameters.ParamByName('patientname').Value:=aSuperArray[i]['��������'].AsString;
        adotemp11.Parameters.ParamByName('sex').Value:=aSuperArray[i]['�����Ա�'].AsString;
        adotemp11.Parameters.ParamByName('age').Value:=aSuperArray[i]['��������'].AsString;
        adotemp11.Parameters.ParamByName('Caseno').Value:=aSuperArray[i]['������'].AsString;
        adotemp11.Parameters.ParamByName('report_date').Value:=RequestDate;
        adotemp11.Parameters.ParamByName('deptname').Value:=aSuperArray[i]['�������'].AsString;
        adotemp11.Parameters.ParamByName('check_doctor').Value:=aSuperArray[i]['����ҽ��'].AsString;
        adotemp11.Parameters.ParamByName('His_Unid').Value:=aSuperArray[i]['ҽ��Ψһ���'].AsString;
        adotemp11.Parameters.ParamByName('Diagnosetype').Value:=YXJB;
        adotemp11.Parameters.ParamByName('flagetype').Value:=SampleType;
        adotemp11.Parameters.ParamByName('typeflagcase').Value:=SampleStatus;
        adotemp11.Parameters.ParamByName('LSH').Value:=lsh;
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
      ExecSQLCmd(AAdoconnstr,'update chk_valu set issure=1 where pkunid='+chk_con_unid+' and pkcombin_id='''+aSuperArrayMX[j]['LIS���������Ŀ����'].AsString+''' and isnull(issure,'''')<>''1'' ');

      adoconn22:=Tadoconnection.Create(nil);
      adoconn22.ConnectionString:=strpas(AAdoconnstr);
      adoconn22.LoginPrompt:=false;

      adotemp22:=Tadoquery.Create(nil);
      adotemp22.Connection:=adoconn22;
      adotemp22.Close;
      adotemp22.SQL.Clear;
      adotemp22.SQL.Text:='select cci.itemid from CombSChkItem csci,combinitem ci,clinicchkitem cci '+
        ' where csci.CombUnid=ci.Unid and cci.unid=csci.ItemUnid and ci.Id='''+aSuperArrayMX[j]['LIS���������Ŀ����'].AsString+''' ';
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
        if '1'<>ScalarSQLCmd(AAdoconnstr,'select top 1 1 from chk_valu where pkunid='+chk_con_unid+' and pkcombin_id='''+aSuperArrayMX[j]['LIS���������Ŀ����'].AsString+''' and itemid='''+adotemp22.FieldByName('itemid').AsString+''' ') then
          ExecSQLCmd(AAdoconnstr,'insert into chk_valu (pkunid,pkcombin_id,itemid,issure) values ('+chk_con_unid+','''+aSuperArrayMX[j]['LIS���������Ŀ����'].AsString+''','''+adotemp22.FieldByName('itemid').AsString+''',1)');

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
