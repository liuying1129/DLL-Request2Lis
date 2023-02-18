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
//该DLL提供函数，用于将JSON格式的检验申请信息导入LIS
//3个输入参数:
//AAdoconnstr:LIS数据库的连接字符串
//ARequestJSON:JSON格式的检验申请单信息
//CurrentWorkGroup:当前工作组.如组合项目未设置默认工作组,申请单将导入当前工作组
//JSON实例:
//  {
//      "检验医嘱": [
//          {
//              "医嘱唯一编号": "10000",
//              "病历号": "101234",
//              "患者姓名": "曹操",
//              "患者性别": "男",
//              "患者年龄": "24岁",
//              "申请日期": "2023-2-6",
//              "申请科室": "内科",
//              "申请医生": "华歆",
//              "医嘱明细": [
//                  {
//                      "联机号": "S0087",
//                      "LIS检验组合项目代码": "06",
//                      "优先级别": "常规",
//                      "样本类型": "血清",
//                      "样本状态": "正常"
//                  },
//                  {
//                      "联机号": "X0013",
//                      "LIS检验组合项目代码": "54",
//                      "优先级别": "常规",
//                      "样本类型": "全血",
//                      "样本状态": "正常"
//                  }
//              ]
//          },
//          {
//              "医嘱唯一编号": "10001",
//              "病历号": "101221",
//              "患者姓名": "关羽",
//              "患者性别": "男",
//              "患者年龄": "25岁",
//              "申请日期": "2023-2-7",
//              "申请科室": "外科",
//              "申请医生": "张飞",
//              "医嘱明细": [
//                  {
//                      "联机号": "S0088",
//                      "LIS检验组合项目代码": "06",
//                      "优先级别": "常规",
//                      "样本类型": "血清",
//                      "样本状态": "正常"
//                  },
//                  {
//                      "联机号": "X0014",
//                      "LIS检验组合项目代码": "54",
//                      "优先级别": "常规",
//                      "样本类型": "全血",
//                      "样本状态": "正常"
//                  }
//              ]
//          }
//      ]
//  }
//上述JSON所有字段必须存在
//值必填的字段：医嘱唯一编号。【医嘱唯一编号】是向HIS返回检验结果的标识,且是程序中子项目插入同一张检验单的判断条件
//JSON中日期时间格式：YYYY-MM-DD hh:nn:ss
//如果【LIS检验组合项目代码】的值在LIS中不存在，则只会导入病人基本信息，不会导入检验项目
//
//2023-02-17本程序已根据工作组、样本类型为依据进行拆单
//是否还要根据子项目【联机字母】进行拆单？观察应用情况再定
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
      MessageDlg('函数ExecSQLCmd失败:'+E.Message+'。错误的SQL:'+ASQL,mtError,[MBOK],0);
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
      MessageDlg('函数ScalarSQLCmd失败:'+E.Message+'。错误的SQL:'+ASQL,mtError,[MBOK],0);
      Qry.Free;
      Conn.Free;
      exit;
    end;
  end;
  Result:=Qry.Fields[0].AsString;
  Qry.Free;
  Conn.Free;
end;

//将医嘱JSON串导入LIS
procedure RequestForm2Lis(const AAdoconnstr,ARequestJSON,CurrentWorkGroup:PChar);stdcall;
var
  adoconn11,adoconn22:Tadoconnection;
  adotemp11,adotemp22:Tadoquery;
  aJson:ISuperObject;
  aSuperArray,aSuperArrayMX: TSuperArray;
  i,j:integer;
  defaultWorkGroup:string;//默认工作组
  defaultSampleType:string;//默认样本类型
  WorkGroup:string;
  SampleType:string;
  chk_con_unid:string;
  YXJB:STRING;//优先级别
  SampleStatus:string;//样本状态
  fs:TFormatSettings;
  RequestDate:TDateTime;//申请日期
  ServerDateTime:TDateTime;
  lsh:string;
begin
  ServerDateTime:=GetServerDate(AAdoconnstr);

  aJson:=SO(ARequestJSON);
  aSuperArray:=aJson['检验医嘱'].AsArray;
  for i:=0 to aSuperArray.Length-1 do
  begin
    aSuperArrayMX:=aSuperArray[i]['医嘱明细'].AsArray;
    for j:=0 to aSuperArrayMX.Length-1 do
    begin
      defaultWorkGroup:=ScalarSQLCmd(AAdoconnstr,'select dept_DfValue from combinitem where Id='''+aSuperArrayMX[j]['LIS检验组合项目代码'].AsString+''' ');
      defaultSampleType:=ScalarSQLCmd(AAdoconnstr,'select specimentype_DfValue from combinitem where Id='''+aSuperArrayMX[j]['LIS检验组合项目代码'].AsString+''' ');

      //如果默认工作组为空,则导入当前工作组
      WorkGroup:=defaultWorkGroup;
      if WorkGroup='' then WorkGroup:=CurrentWorkGroup;

      //如果JSON中样本类型为空,则取默认样本类型
      SampleType:=aSuperArrayMX[j]['样本类型'].AsString;
      if SampleType='' then SampleType:=defaultSampleType;

      YXJB:=aSuperArrayMX[j]['优先级别'].AsString;
      if YXJB='' then YXJB:='常规';

      SampleStatus:=aSuperArrayMX[j]['样本状态'].AsString;
      if SampleStatus='' then SampleStatus:='正常';

      fs.DateSeparator:='-';
      fs.TimeSeparator:=':';
      fs.ShortDateFormat:='YYYY-MM-DD hh:nn:ss';
      RequestDate:=StrtoDateTimeDef(aSuperArray[i]['申请日期'].AsString,ServerDateTime,fs);
      if  RequestDate<2 then ReplaceDate(RequestDate,ServerDateTime);//表示1899-12-30,没有给日期赋值
      if (HourOf(RequestDate)=0) and (MinuteOf(RequestDate)=0) and (SecondOf(RequestDate)=0) then ReplaceTime(RequestDate,ServerDateTime);//表示没有给时间赋值

      chk_con_unid:=ScalarSQLCmd(AAdoconnstr,'select top 1 unid from chk_con cc where cc.combin_id='''+WorkGroup+''' and cc.His_Unid='''+aSuperArray[i]['医嘱唯一编号'].AsString+''' and cc.flagetype='''+SampleType+''' and isnull(report_doctor,'''')='''' ');
      if chk_con_unid='' then//存在工作组、医嘱唯一编号、样本类型相同,且未审核的检验单,则在此检验单上新增明细.否则就新增一条检验单
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
        adotemp11.Parameters.ParamByName('checkid').Value:=aSuperArrayMX[j]['联机号'].AsString;
        adotemp11.Parameters.ParamByName('patientname').Value:=aSuperArray[i]['患者姓名'].AsString;
        adotemp11.Parameters.ParamByName('sex').Value:=aSuperArray[i]['患者性别'].AsString;
        adotemp11.Parameters.ParamByName('age').Value:=aSuperArray[i]['患者年龄'].AsString;
        adotemp11.Parameters.ParamByName('Caseno').Value:=aSuperArray[i]['病历号'].AsString;
        adotemp11.Parameters.ParamByName('report_date').Value:=RequestDate;
        adotemp11.Parameters.ParamByName('deptname').Value:=aSuperArray[i]['申请科室'].AsString;
        adotemp11.Parameters.ParamByName('check_doctor').Value:=aSuperArray[i]['申请医生'].AsString;
        adotemp11.Parameters.ParamByName('His_Unid').Value:=aSuperArray[i]['医嘱唯一编号'].AsString;
        adotemp11.Parameters.ParamByName('Diagnosetype').Value:=YXJB;
        adotemp11.Parameters.ParamByName('flagetype').Value:=SampleType;
        adotemp11.Parameters.ParamByName('typeflagcase').Value:=SampleStatus;
        adotemp11.Parameters.ParamByName('LSH').Value:=lsh;
        Try
          adotemp11.Open;
        except
          on E:Exception do
          begin
            MessageDlg('插入病人信息失败:'+E.Message,mtError,[MBOK],0);
            adotemp11.Free;
            adoconn11.Free;
            exit;
          end;
        end;
        chk_con_unid:=adotemp11.fieldbyname('Insert_Identity').AsString;
        adotemp11.Free;
        adoconn11.Free;
      end;

      //插入明细begin
      ExecSQLCmd(AAdoconnstr,'update chk_valu set issure=1 where pkunid='+chk_con_unid+' and pkcombin_id='''+aSuperArrayMX[j]['LIS检验组合项目代码'].AsString+''' and isnull(issure,'''')<>''1'' ');

      adoconn22:=Tadoconnection.Create(nil);
      adoconn22.ConnectionString:=strpas(AAdoconnstr);
      adoconn22.LoginPrompt:=false;

      adotemp22:=Tadoquery.Create(nil);
      adotemp22.Connection:=adoconn22;
      adotemp22.Close;
      adotemp22.SQL.Clear;
      adotemp22.SQL.Text:='select cci.itemid from CombSChkItem csci,combinitem ci,clinicchkitem cci '+
        ' where csci.CombUnid=ci.Unid and cci.unid=csci.ItemUnid and ci.Id='''+aSuperArrayMX[j]['LIS检验组合项目代码'].AsString+''' ';
      Try
        adotemp22.Open;
      except
        on E:Exception do
        begin
          MessageDlg('获取指定组合项目的子项目失败:'+E.Message,mtError,[MBOK],0);
          adotemp22.Free;
          adoconn22.Free;
          exit;
        end;
      end;
      while not adotemp22.Eof do
      begin
        if '1'<>ScalarSQLCmd(AAdoconnstr,'select top 1 1 from chk_valu where pkunid='+chk_con_unid+' and pkcombin_id='''+aSuperArrayMX[j]['LIS检验组合项目代码'].AsString+''' and itemid='''+adotemp22.FieldByName('itemid').AsString+''' ') then
          ExecSQLCmd(AAdoconnstr,'insert into chk_valu (pkunid,pkcombin_id,itemid,issure) values ('+chk_con_unid+','''+aSuperArrayMX[j]['LIS检验组合项目代码'].AsString+''','''+adotemp22.FieldByName('itemid').AsString+''',1)');

        adotemp22.Next;
      end;
      adotemp22.Free;
      adoconn22.Free;

      //Data2Lis传入结果时也会调用，故此处先注释
      //addOrEditCalcItem(pchar(LisConn),pchar(s2),checkunid);//增加计算项目
      //addOrEditCalcValu(pchar(LisConn),checkunid,false,'');//更新计算项目
      //插入明细end
    end;
  end;
end;

exports
  RequestForm2Lis;

begin
end.
