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
//      "JSON数据源":"HIS",
//      "检验医嘱": [
//          {
//              "申请单编号": "10000",
//              "病历号": "101234",
//              "患者姓名": "曹操",
//              "患者性别": "男",
//              "患者年龄": "24岁",
//              "申请日期": "2023-2-6",
//              "申请科室": "内科",
//              "申请医生": "华歆",
//              "床号":"",
//              "临床诊断":"",
//              "备注":"",
//              "所属公司":"",
//              "所属部门":"",
//              "工种":"",
//              "工号":"",
//              "婚否":"",
//              "籍贯":"",
//              "住址":"",
//              "电话":"",
//              "外部系统唯一编号":"",
//              "医嘱明细": [
//                  {
//                      "联机号": "S0087",
//                      "LIS组合项目代码": "06",
//                      "条码号": "12345",
//                      "优先级别": "常规",
//                      "样本类型": "血清",
//                      "样本状态": "正常"
//                  },
//                  {
//                      "联机号": "X0013",
//                      "LIS组合项目代码": "54",
//                      "条码号": "12346",
//                      "优先级别": "常规",
//                      "样本类型": "全血",
//                      "样本状态": "正常"
//                  }
//              ]
//          },
//          {
//              "申请单编号": "10001",
//              "病历号": "101221",
//              "患者姓名": "关羽",
//              "患者性别": "男",
//              "患者年龄": "25岁",
//              "申请日期": "2023-2-7",
//              "申请科室": "外科",
//              "申请医生": "张飞",
//              "床号":"",
//              "临床诊断":"",
//              "备注":"",
//              "所属公司":"",
//              "所属部门":"",
//              "工种":"",
//              "工号":"",
//              "婚否":"",
//              "籍贯":"",
//              "住址":"",
//              "电话":"",
//              "外部系统唯一编号":"",
//              "医嘱明细": [
//                  {
//                      "联机号": "S0088",
//                      "LIS组合项目代码": "06",
//                      "条码号": "12347",
//                      "优先级别": "常规",
//                      "样本类型": "血清",
//                      "样本状态": "正常"
//                  },
//                  {
//                      "联机号": "X0014",
//                      "LIS组合项目代码": "54",
//                      "条码号": "12348",
//                      "优先级别": "常规",
//                      "样本类型": "全血",
//                      "样本状态": "正常"
//                  }
//              ]
//          }
//      ]
//  }
//JSON必须存在的key：JSON数据源、检验医嘱、条码号(特别的，如果JSON数据源的值为Excel，该key可以不存在)、医嘱明细
//【JSON数据源】值必填：HIS、Excel
//【条码号】：当【JSON数据源】值为HIS时,【条码号】是程序中子项目插入同一张检验单的判断条件
//【外部系统唯一编号】: 当【JSON数据源】值为HIS时,HIS/PEIS等外部系统可用此编号关联受检者与检验结果.此编号有可能是体检号,也有可能是HIS表示此次看病的看病号
//如果【LIS组合项目代码】的值在LIS中不存在，则仅会导入病人基本信息，不会导入检验项目
//如果希望仅导入病人基本信息,则需要保证【医嘱明细】至少有一条记录,哪怕是一条无效数据的记录
//JSON中日期时间格式：YYYY-MM-DD hh:nn:ss
//
//2023-02-17本程序已根据工作组、样本类型为依据进行拆单
//是否还要根据子项目【联机字母】进行拆单？观察应用情况再定
//
//将JSON显示为脑图的网站:https://jsoncrack.com/editor
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
  lsh:string;//流水号
  pkcombin_id:String;//LIS组合项目代码
  RequestDateStr:String;//申请日期
  checkid:String;//联机号
  patientname:String;//患者姓名
  sex:String;//患者性别
  age:String;//患者年龄
  Caseno:String;//病历号
  deptname:String;//申请科室
  check_doctor:String;//申请医生
  bedno:String;//床号
  diagnose:String;//临床诊断
  issure:String;//备注
  WorkCompany:String;//所属公司
  WorkDepartment:String;//所属部门
  WorkCategory:String;//工种
  WorkID:String;//工号
  ifMarry:String;//婚否
  OldAddress:String;//籍贯
  Address:String;//住址
  Telephone:String;//电话
  DNH:String;//申请单编号(HIS)
  His_Unid:String;//外部系统唯一编号(HIS)
begin
  ServerDateTime:=GetServerDate(AAdoconnstr);

  aJson:=SO(ARequestJSON);
  if not aJson.AsObject.Exists('JSON数据源') then exit;//判断key是否存在的另一种写法:if aJson['JSON数据源']=nil then exit;
  if not aJson.AsObject.Exists('检验医嘱') then exit;
  
  aSuperArray:=aJson['检验医嘱'].AsArray;
  for i:=0 to aSuperArray.Length-1 do
  begin
    if not aSuperArray[i].AsObject.Exists('医嘱明细') then continue;

    aSuperArrayMX:=aSuperArray[i]['医嘱明细'].AsArray;
    for j:=0 to aSuperArrayMX.Length-1 do
    begin
      if ('Excel'<>aJson.S['JSON数据源'])and(not aSuperArrayMX[j].AsObject.Exists('条码号')) then continue;

      if aSuperArrayMX[j].AsObject.Exists('LIS组合项目代码') then pkcombin_id:=aSuperArrayMX[j]['LIS组合项目代码'].AsString else pkcombin_id:=''; 
      if pkcombin_id='' then pkcombin_id:='不存在的组合项目代码';
        
      defaultWorkGroup:=ScalarSQLCmd(AAdoconnstr,'select dept_DfValue from combinitem where Id='''+pkcombin_id+''' ');
      defaultSampleType:=ScalarSQLCmd(AAdoconnstr,'select specimentype_DfValue from combinitem where Id='''+pkcombin_id+''' ');

      //如果默认工作组为空,则导入当前工作组
      WorkGroup:=defaultWorkGroup;
      if WorkGroup='' then WorkGroup:=CurrentWorkGroup;
      if 'Excel'=aJson['JSON数据源'].AsString then WorkGroup:=CurrentWorkGroup;

      //如果JSON中样本类型为空,则取默认样本类型
      if aSuperArrayMX[j].AsObject.Exists('样本类型') then SampleType:=aSuperArrayMX[j]['样本类型'].AsString else SampleType:=''; 
      if (SampleType='')and('Excel'<>aJson['JSON数据源'].AsString) then SampleType:=defaultSampleType;

      if aSuperArrayMX[j].AsObject.Exists('优先级别') then YXJB:=aSuperArrayMX[j]['优先级别'].AsString else YXJB:=''; 
      if YXJB='' then YXJB:='常规';

      if aSuperArrayMX[j].AsObject.Exists('样本状态') then SampleStatus:=aSuperArrayMX[j]['样本状态'].AsString else SampleStatus:=''; 
      if SampleStatus='' then SampleStatus:='正常';

      fs.DateSeparator:='-';
      fs.TimeSeparator:=':';
      fs.ShortDateFormat:='YYYY-MM-DD hh:nn:ss';
      if aSuperArray[i].AsObject.Exists('申请日期') then RequestDateStr:=aSuperArray[i]['申请日期'].AsString else RequestDateStr:='';
      RequestDate:=StrtoDateTimeDef(RequestDateStr,ServerDateTime,fs);
      if  RequestDate<2 then ReplaceDate(RequestDate,ServerDateTime);//表示1899-12-30,没有给日期赋值
      if (HourOf(RequestDate)=0) and (MinuteOf(RequestDate)=0) and (SecondOf(RequestDate)=0) then ReplaceTime(RequestDate,ServerDateTime);//表示没有给时间赋值

      if aSuperArrayMX[j].AsObject.Exists('联机号') then checkid:=aSuperArrayMX[j]['联机号'].AsString else checkid:='';
      if aSuperArray[i].AsObject.Exists('患者姓名') then patientname:=aSuperArray[i]['患者姓名'].AsString else patientname:='';
      if aSuperArray[i].AsObject.Exists('患者性别') then sex:=aSuperArray[i]['患者性别'].AsString else sex:=''; 
      if aSuperArray[i].AsObject.Exists('患者年龄') then age:=aSuperArray[i]['患者年龄'].AsString else age:='';
      if aSuperArray[i].AsObject.Exists('病历号') then Caseno:=aSuperArray[i]['病历号'].AsString else Caseno:='';
      if aSuperArray[i].AsObject.Exists('申请科室') then deptname:=aSuperArray[i]['申请科室'].AsString else deptname:='';
      if aSuperArray[i].AsObject.Exists('申请医生') then check_doctor:=aSuperArray[i]['申请医生'].AsString else check_doctor:='';
      if aSuperArray[i].AsObject.Exists('床号') then bedno:=aSuperArray[i]['床号'].AsString else bedno:='';
      if aSuperArray[i].AsObject.Exists('临床诊断') then diagnose:=aSuperArray[i]['临床诊断'].AsString else diagnose:='';
      if aSuperArray[i].AsObject.Exists('备注') then issure:=aSuperArray[i]['备注'].AsString else issure:='';
      if aSuperArray[i].AsObject.Exists('所属公司') then WorkCompany:=aSuperArray[i]['所属公司'].AsString else WorkCompany:='';
      if aSuperArray[i].AsObject.Exists('所属部门') then WorkDepartment:=aSuperArray[i]['所属部门'].AsString else WorkDepartment:='';
      if aSuperArray[i].AsObject.Exists('工种') then WorkCategory:=aSuperArray[i]['工种'].AsString else WorkCategory:='';
      if aSuperArray[i].AsObject.Exists('工号') then WorkID:=aSuperArray[i]['工号'].AsString else WorkID:='';
      if aSuperArray[i].AsObject.Exists('婚否') then ifMarry:=aSuperArray[i]['婚否'].AsString else ifMarry:='';
      if aSuperArray[i].AsObject.Exists('籍贯') then OldAddress:=aSuperArray[i]['籍贯'].AsString else OldAddress:='';
      if aSuperArray[i].AsObject.Exists('住址') then Address:=aSuperArray[i]['住址'].AsString else Address:='';
      if aSuperArray[i].AsObject.Exists('电话') then Telephone:=aSuperArray[i]['电话'].AsString else Telephone:='';
      if aSuperArray[i].AsObject.Exists('外部系统唯一编号') then His_Unid:=aSuperArray[i]['外部系统唯一编号'].AsString else His_Unid:='';
      if aSuperArray[i].AsObject.Exists('申请单编号') then DNH:=aSuperArray[i]['申请单编号'].AsString else DNH:='';

      if 'Excel'=aJson['JSON数据源'].AsString then chk_con_unid:=ScalarSQLCmd(AAdoconnstr,'select top 1 unid from chk_con where patientname='''+patientname+''' AND sex='''+sex+''' AND age='''+age+''' AND combin_id='''+WorkGroup+''' and isnull(report_doctor,'''')='''' ')
        else chk_con_unid:=ScalarSQLCmd(AAdoconnstr,'select top 1 unid from chk_con cc where cc.combin_id='''+WorkGroup+''' and cc.TjJianYan='''+aSuperArrayMX[j]['条码号'].AsString+''' and cc.flagetype='''+SampleType+''' and isnull(report_doctor,'''')='''' ');
        
      if chk_con_unid='' then//存在工作组、条码号、样本类型相同,且未审核的检验单,则在此检验单上新增明细.否则就新增一条检验单
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
        if 'Excel'=aJson['JSON数据源'].AsString then
          adotemp11.Parameters.ParamByName('TjJianYan').Value:=''
        else
          adotemp11.Parameters.ParamByName('TjJianYan').Value:=aSuperArrayMX[j]['条码号'].AsString;
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
          MessageDlg('获取指定组合项目的子项目失败:'+E.Message,mtError,[MBOK],0);
          adotemp22.Free;
          adoconn22.Free;
          exit;
        end;
      end;
      while not adotemp22.Eof do
      begin
        if '1'<>ScalarSQLCmd(AAdoconnstr,'select top 1 1 from chk_valu where pkunid='+chk_con_unid+' and pkcombin_id='''+aSuperArrayMX[j]['LIS组合项目代码'].AsString+''' and itemid='''+adotemp22.FieldByName('itemid').AsString+''' ') then
          ExecSQLCmd(AAdoconnstr,'insert into chk_valu (pkunid,pkcombin_id,itemid,issure) values ('+chk_con_unid+','''+aSuperArrayMX[j]['LIS组合项目代码'].AsString+''','''+adotemp22.FieldByName('itemid').AsString+''',1)');

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
