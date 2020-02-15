using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Security;
using System.Text;
using System.Threading.Tasks;
using Microsoft.VisualBasic;
using System.Web;
using System.Web.SessionState;
using System.Data;
using LeaScentMark.AutoCode;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json.Converters;
using Ionic.Zip;



public class UserInit : IHttpHandler, IRequiresSessionState
{
    public void ProcessRequest(HttpContext context)
    {
        string jsonResult = "";

        string requestType = context.Request.Params["HTTP_X_REQUESTED_WITH"];   // "XMLHttpRequest"	String

        if (string.IsNullOrEmpty(context.Session["LoginUserID"]) && requestType.Equals("XMLHttpRequest", StringComparison.CurrentCultureIgnoreCase))
        {
            // response.setHeader("sessionstatus", "timeout")
            context.Response.AddHeader("sessionstatus", "timeout");
            context.Response.ContentType = "application/json";
            context.Response.Write("");
            context.Response.End();
            return;
        }

        var method = context.Request["method"] ?? "";
        if (!string.IsNullOrEmpty(method))
        {
            switch (method)
            {
                case "StuGetScoreInfoByTestID":
                    {
                        jsonResult = StuGetScoreInfoByTestID(context);
                        break;
                        break;
                    }

                case "NGenStuReportByTestID":
                    {
                        jsonResult = NGenStuReportByTestID(context);
                        break;
                        break;
                    }

                case "GetUserLicInfo":
                    {
                        jsonResult = GetUserLicInfo(context);
                        break;
                        break;
                    }

                case "GetReportUnitNest":
                    {
                        jsonResult = GetReportUnitNest(context);
                        break;
                        break;
                    }

                case "GetUserLicTestList":
                    {
                        jsonResult = GetUserLicTestList(context);
                        break;
                        break;
                    }

                case "GetClassStuLIDScore":
                    {
                        jsonResult = GetClassStuLIDScore(context);
                        break;
                        break;
                    }

                case "GenHtmlReport":
                    {
                        jsonResult = GenHtmlReport(context);
                        break;
                        break;
                    }

                case "GenHtmlReportByPage":
                    {
                        jsonResult = GenHtmlReportByPage(context);
                        break;
                        break;
                    }

                case "GenStuAllTestInfo":
                    {
                        jsonResult = GenStuAllTestInfo(context);
                        break;
                        break;
                    }

                case "GetQuesMedia":
                    {
                        jsonResult = GetQuesMedia(context);
                        break;
                        break;
                    }

                case "GetClassTestLessonReport":
                    {
                        jsonResult = GetClassTestLessonReport(context);
                        break;
                        break;
                    }

                case "ChangUserPW":
                    {
                        jsonResult = ChangUserPW(context);
                        break;
                        break;
                    }

                case "DownLoadFiles":
                    {
                        DownLoadFiles(context);
                        return;
                    }

                case "GetStuPaperURL":
                    {
                        jsonResult = GetStuPaperURL(context);
                        break;
                        break;
                    }

                case "GetClassLessonTestList":
                    {
                        jsonResult = GetClassLessonTestList(context);
                        break;
                        break;
                    }

                case "LogOffSys":
                    {
                        jsonResult = LogOffSys(context);
                        break;
                        break;
                    }

                case "ExportReport"  // 下载报表
         :
                    {
                        ExportReport(context);
                        return;
                    }
            }
        }


        context.Response.ContentType = "application/json";
        context.Response.Write(jsonResult);
        context.Response.End();
    }

    public bool IsReusable
    {
        get
        {
            return false;
        }
    }


    public string GetUserLicInfo(HttpContext context)
    {
        string sUserID = context.Session["LoginUserID"].ToString();

        int iRequestGradeID = 0;

        if (context.Request.QueryString["GradeID"] != null)
            int.TryParse(context.Request.QueryString["GradeID"].ToString(), ref iRequestGradeID);

        IDataParameter[] parm = new IDataParameter[2] { };

        parm[0] = DBHelp.CreateParameter("@UserID", sUserID);
        parm[1] = DBHelp.CreateParameter("@RequestGradeID", iRequestGradeID);


        DataSet ds = DBHelp.GetDataSet("pr_GetUserLicInfo", parm);

        int dsTableCount = ds.Tables.Count;

        TestInfo tTestInfo = new TestInfo();

        DataRow dr;

        dr = ds.Tables[0].Rows[0];

        tTestInfo.CurrentGradeID = dr["GradeID"].ToString();
        tTestInfo.TestID = dr["TestID"].ToString();
        tTestInfo.TestName = dr["TestName"].ToString();
        tTestInfo.UserName = dr["UserName"].ToString();
        tTestInfo.RoleName = dr["RoleName"].ToString();
        tTestInfo.FilterName = dr["FilterName"].ToString();

        foreach (var dr in ds.Tables[1].Rows)
        {
            LessonInfo tLessonInfo = new LessonInfo();
            tLessonInfo.LessonID = dr["LessonID"].ToString();
            tLessonInfo.LessonName = dr["LessonID"].ToString();
            tLessonInfo.MultiLesson = dr["MultiLesson"].ToString();

            tTestInfo.LessonList.Add(tLessonInfo);
        }

        if (dsTableCount > 2 && iRequestGradeID == 0)
        {
            foreach (var dr in ds.Tables[2].Rows)
            {
                GradeInfo tGradeInfo = new GradeInfo();
                tGradeInfo.GradeID = dr["GradeID"].ToString();
                tGradeInfo.GradeName = dr["GradeName"].ToString();

                tTestInfo.GradeList.Add(tGradeInfo);
            }
        }

        if (dsTableCount > 3)
        {
            foreach (var dr in ds.Tables[3].Rows)
            {
                ReportInfo tReportInfo = new ReportInfo();
                tReportInfo.ReportID = dr["ReportID"].ToString();
                tReportInfo.ReportName = dr["ReportName"].ToString();
                tReportInfo.MultiLesson = dr["MultiLesson"].ToString();
                tReportInfo.ReportURL = "";
                tTestInfo.ReportList.Add(tReportInfo);
            }
        }

        if (dsTableCount > 4)
        {
            foreach (var dr in ds.Tables[4].Rows)
            {
                MenuInfo tMenuInfo = new MenuInfo();
                tMenuInfo.MenuLevelID = dr["LicLevel"].ToString();
                tMenuInfo.MenuName = dr["mName"].ToString();
                tMenuInfo.MenuURL = dr["mURL"].ToString();
                tTestInfo.MenuList.Add(tMenuInfo);
            }
        }

        GetUserLicInfo = JsonConvert.SerializeObject(tTestInfo);
    }


    public string GetReportUnitNest(HttpContext context)
    {
        string sUserID = context.Session["LoginUserID"].ToString();
        string iTestID = ""; // context.Session("TestID").ToString()
        string iLevel = "";

        IDataParameter[] parm = new IDataParameter[3] { };

        if (context.Request.QueryString["TestID"] != null)
            iTestID = context.Request.QueryString["TestID"].ToString();

        if (context.Request.QueryString["LevelID"] != null)
            iLevel = context.Request.QueryString["LevelID"].ToString();

        parm[0] = DBHelp.CreateParameter("@TestID", iTestID);
        parm[1] = DBHelp.CreateParameter("@UserID", sUserID);
        parm[2] = DBHelp.CreateParameter("@DisplayLevel", iLevel);


        DataSet ds = DBHelp.GetDataSet("pr_NReportGetLicUnit", parm);
        ds.Tables[0].TableName = "UnitType";
        ds.Tables[1].TableName = "ParentUnit";
        ds.Tables[2].TableName = "ChildUnit";

        // 必须带上False 否则 不能启用此约束 因为不是所有的值都具有相应的父值
        ds.Relations.Add("UnitChild", ds.Tables["ParentUnit"].Columns["ChildID"], ds.Tables["ChildUnit"].Columns["ParentID"], false);

        DataRow rowParent, rowChild;



        ReportUnit m_UnitInfo = new ReportUnit();

        // Dim m_UnitLevel = New List(Of ReportUnitLevel)()

        foreach (var rowParent in ds.Tables["UnitType"].Rows)
        {
            ReportUnitLevel sUnitType = new ReportUnitLevel();
            sUnitType.UnitTypeID = rowParent["UnitType"].ToString();
            sUnitType.UnitTypeName = rowParent["UnitTypeName"].ToString();
            m_UnitInfo.UnitLevel.Add(sUnitType);
        }

        // Dim m_UnitNest = New List(Of ReportUnitNest)()

        foreach (var rowParent in ds.Tables["ParentUnit"].Rows)
        {
            ReportUnitNest sUnit = new ReportUnitNest();

            sUnit.UnitID = rowParent["UnitName"].ToString();
            sUnit.UnitName = rowParent["UnitName"].ToString();

            foreach (var rowChild in rowParent.GetChildRows("UnitChild", DataRowVersion.Original))
            {
                ReportUnitChild sUnitChild = new ReportUnitChild();
                sUnitChild.UnitID = rowChild["UnitID"].ToString();
                sUnitChild.UnitName = rowChild["UnitName"].ToString();
                sUnitChild.UnitLevel = rowChild["UnitType"].ToString();

                sUnit.ChildUnit.Add(sUnitChild);
            }
            m_UnitInfo.UnitInfo.Add(sUnit);
        }


        GetReportUnitNest = JsonConvert.SerializeObject(m_UnitInfo);
    }

    public string GetUserLicTestList(HttpContext context)
    {
        string sGradeID = ""; // context.Session("LoginUserID").ToString()
        string sLessonID = "";

        int iRequestGradeID = 0;

        if (context.Request.QueryString["GradeID"] != null)
            sGradeID = context.Request.QueryString["GradeID"].ToString();
        if (context.Request.QueryString["LessonID"] != null)
            sLessonID = context.Request.QueryString["LessonID"].ToString();
        IDataParameter[] parm = new IDataParameter[2] { };
        parm[0] = DBHelp.CreateParameter("@GradeID", sGradeID);
        parm[1] = DBHelp.CreateParameter("@LessonID", sLessonID);

        DataSet ds = DBHelp.GetDataSet("pr_GetTestListByGLID", parm);

        ds.Tables[0].TableName = "TestList";
        ds.Tables[1].TableName = "TestType";

        GetUserLicTestList = JsonConvert.SerializeObject(ds);
    }


    public string GetClassStuLIDScore(HttpContext context)
    {
        string sTestID;
        string sLessonID;
        string sStatUnitID;
        string sUserID = context.Session["LoginUserID"].ToString();

        if (context.Request.QueryString["resData"] != null)
        {
            string[] sPara = context.Request.QueryString["resData"].ToString().Split("|");
            sStatUnitID = sPara[0];
            sTestID = sPara[1];
            sLessonID = sPara[2];
        }

        // If context.Request.QueryString("TestID") IsNot Nothing Then
        // sTestID = context.Request.QueryString("TestID").ToString()
        // End If
        // If context.Request.QueryString("LessonID") IsNot Nothing Then
        // sLessonID = context.Request.QueryString("LessonID").ToString()
        // End If

        // If context.Request.QueryString("StatUnitID") IsNot Nothing Then
        // sStatUnitID = context.Request.QueryString("StatUnitID").ToString()
        // End If

        IDataParameter[] parm = new IDataParameter[4] { };
        parm[0] = DBHelp.CreateParameter("@UserID", sUserID);
        parm[1] = DBHelp.CreateParameter("@TestID", sTestID);
        parm[2] = DBHelp.CreateParameter("@LessonID", sLessonID);
        parm[3] = DBHelp.CreateParameter("@StatUnitID", sStatUnitID);

        DataSet ds = DBHelp.GetDataSet("pr_GetClassStuLIDScore", parm);

        GetClassStuLIDScore = JsonConvert.SerializeObject(ds);
    }


    public string StuGetScoreInfoByTestID(HttpContext context)
    {
        string sTestID;   // context.Session("LoginUserID").ToString()
        string sXHID;
        string sStatUnitID;

        if (context.Request.QueryString["resData"] != null)
        {
            string[] sPara = context.Request.QueryString["resData"].ToString().Split("|");
            sXHID = sPara[0];
            sTestID = sPara[1];
        }

        // If context.Request.QueryString("TestID") IsNot Nothing Then
        // sTestID = context.Request.QueryString("TestID").ToString()
        // End If
        // If context.Request.QueryString("XHID") IsNot Nothing Then
        // sXHID = context.Request.QueryString("XHID").ToString()
        // End If

        // If context.Request.QueryString("StatUnit") IsNot Nothing Then
        // sStatUnitID = context.Request.QueryString("StatUnit").ToString()
        // End If

        IDataParameter[] parm = new IDataParameter[2] { };

        parm[0] = DBHelp.CreateParameter("@TestID", sTestID);
        parm[1] = DBHelp.CreateParameter("@XHID", sXHID);

        DataSet ds = DBHelp.GetDataSet("pr_StuGetScoreInfoByTestID", parm);

        ds.Tables[0].TableName = "StuName";
        ds.Tables[1].TableName = "StuList";
        ds.Tables[2].TableName = "TestList";

        ds.Relations.Add("LessonList", ds.Tables[2].Columns["TestID"], ds.Tables[3].Columns["TestID"], false);

        JObject jItem;
        JArray jTableData;
        JObject jRowData;
        int iCounter;
        DataRow rowParent;
        DataRow rowChild;

        jItem = new JObject();

        for (iCounter = 0; iCounter <= 1; iCounter++)
        {
            jTableData = new JArray();

            foreach (var rowChild in ds.Tables[iCounter].Rows)
            {
                jRowData = new JObject();

                foreach (DataColumn col in ds.Tables[iCounter].Columns)
                    jRowData.Add(new JProperty(col.ColumnName, rowChild[col.ColumnName]));

                jTableData.Add(jRowData);
            }

            jItem.Add(new JProperty(ds.Tables[iCounter].TableName, jTableData));
        }

        jTableData = new JArray();

        JArray jChildData;
        JObject jRowDataParent;

        foreach (var rowParent in ds.Tables[2].Rows)
        {
            jRowDataParent = new JObject();

            jRowDataParent.Add(new JProperty("TestTitle", rowParent["TestTitle"]));
            jRowDataParent.Add(new JProperty("TestDate", rowParent["TestDate"]));


            jChildData = new JArray();

            foreach (var rowChild in rowParent.GetChildRows("LessonList", DataRowVersion.Original))
            {
                jRowData = new JObject();
                foreach (DataColumn col in ds.Tables[3].Columns)
                    jRowData.Add(new JProperty(col.ColumnName, rowChild[col.ColumnName]));
                jChildData.Add(jRowData);
            }


            jRowDataParent.Add(new JProperty("LessonList", jChildData));

            jTableData.Add(jRowDataParent);
        }

        jItem.Add(new JProperty("TestList", jTableData));

        StuGetScoreInfoByTestID = jItem.ToString();
    }


    public string NGenStuReportByTestID(HttpContext context)
    {
        string sTestID;   // context.Session("LoginUserID").ToString()
        string sXHID;
        string sLessonID;



        if (context.Request.QueryString["TestID"] != null)
            sTestID = context.Request.QueryString["TestID"].ToString();

        int iUserType;
        int.TryParse(context.Session["UserType"], ref iUserType);

        if (iUserType == -1)
            sXHID = context.Session["LoginUserID"].ToString();
        else if (context.Request.QueryString["XHID"] != null)
            sXHID = context.Request.QueryString["XHID"].ToString();






        if (context.Request.QueryString["LessonID"] != null)
            sLessonID = context.Request.QueryString["LessonID"].ToString();

        IDataParameter[] parm = new IDataParameter[3] { };

        parm[0] = DBHelp.CreateParameter("@TestID", sTestID);
        parm[1] = DBHelp.CreateParameter("@LessonID", sLessonID);
        parm[2] = DBHelp.CreateParameter("@XHID", sXHID);

        DataSet ds = DBHelp.GetDataSet("pr_NGenStuReportByTestID", parm);

        JObject jItem = new JObject();


        JArray jTableData = new JArray();
        JObject jRowData = new JObject();
        foreach (DataRow dr in ds.Tables[0].Rows)
        {
            jRowData = new JObject();
            foreach (DataColumn col in ds.Tables[0].Columns)
                jRowData.Add(new JProperty(col.ColumnName, dr[col.ColumnName]));
            jTableData.Add(jRowData);
        }

        jItem.Add(new JProperty("BData", jTableData));



        JArray jSeries = new JArray();

        JArray jCategory = new JArray();
        JArray jSeriesDataPerson = new JArray();
        JArray jSeriesDataClass = new JArray();
        JArray jSeriesDataGrade = new JArray();
        JObject jSeriesItem = new JObject();

        JArray jDataXL = new JArray();
        // 平稳度----------------
        foreach (DataRow row in ds.Tables[1].Rows)
        {
            jDataXL = new JArray();
            jDataXL.Add(row[0]);
            jDataXL.Add(row[1]);
            jSeriesDataClass.Add(jDataXL);
        }


        jSeriesItem.Add(new JProperty("name", "其他同学"));
        jSeriesItem.Add(new JProperty("color", "#C0C0C0"));
        jSeriesItem.Add(new JProperty("data", jSeriesDataClass));
        jSeries.Add(jSeriesItem);

        jSeriesItem = new JObject();

        jSeriesDataPerson = new JArray();
        if (ds.Tables[2].Rows.Count > 0)
        {
            jDataXL = new JArray();
            jDataXL.Add(ds.Tables[2].Rows[0][0]);
            jDataXL.Add(ds.Tables[2].Rows[0][1]);
            jSeriesDataPerson.Add(jDataXL);
            // jSeriesDataPerson.Add("[" + ds.Tables(2).Rows(0)(1).ToString() + "," + ds.Tables(2).Rows(0)(1).ToString() + "]")
            jItem.Add(new JProperty("SPInfo", ds.Tables[2].Rows[0]["Desp"]));
            jItem.Add(new JProperty("SPAtten", Interaction.IIf(ds.Tables[2].Rows[0]["Attention"] == "0", "平稳", "需注意")));
        }
        else
        {
            jSeriesDataPerson.Add("[0,0]");
            jItem.Add(new JProperty("SPInfo", ""));
            jItem.Add(new JProperty("SPAtten", ""));
        }
        jSeriesItem.Add(new JProperty("name", "本人"));
        jSeriesItem.Add(new JProperty("color", "#0C6"));
        jSeriesItem.Add(new JProperty("data", jSeriesDataPerson));
        jSeries.Add(jSeriesItem);

        JObject jChartInfo = new JObject();

        jChartInfo.Add(new JProperty("Category", jCategory));
        jChartInfo.Add(new JProperty("Series", jSeries));


        jItem.Add(new JProperty("PWD", jChartInfo));
        // ----------------平稳度-  

        // 知识点-------------------
        jCategory = new JArray();
        jSeriesDataPerson = new JArray();
        jSeriesDataClass = new JArray();
        jSeriesDataGrade = new JArray();
        jSeries = new JArray();

        int iCounter;
        // For iCounter = 1 To ds.Tables(3).Columns.Count
        // jCategory.Add(iCounter.ToString())
        // Next

        iCounter = 1;
        string KBInfo = "";
        string KHInfo = "";

        foreach (DataRow row in ds.Tables[3].Rows)
        {
            jCategory.Add(row["RID"]);
            // jCategory.Add(row("KnowName"))
            jSeriesDataPerson.Add(row["PDV"]);
            jSeriesDataClass.Add(row["BDV"]);
            // jSeriesDataGrade.Add(row("GDV"))
            if (row["DDVF"] < 0)
                KBInfo = KBInfo + "<li><span style='color:#f00;font-weight:700;'>K" + iCounter.ToString() + "</span>：" + row["KnowName"] + "</li>";
            if (row["DDVF"] > 0)
                KHInfo = KHInfo + "<li><span style='color:#f00;font-weight:700;'>K" + iCounter.ToString() + "</span>：" + row["KnowName"] + "</li>";

            iCounter += 1;
        }

        // If jSeriesDataPerson.Count = 0 Then
        // jCategory.Add(0)
        // jSeriesDataPerson.Add("0")
        // jSeriesDataClass.Add("0")
        // End If

        jItem.Add(new JProperty("KBInfo", KBInfo));
        jItem.Add(new JProperty("KHInfo", KHInfo));

        jSeriesItem = new JObject();
        jSeriesItem.Add(new JProperty("type", "line"));
        jSeriesItem.Add(new JProperty("name", "班级"));
        jSeriesItem.Add(new JProperty("color", "#ff0000"));
        jSeriesItem.Add(new JProperty("data", jSeriesDataClass));
        jSeries.Add(jSeriesItem);

        jSeriesItem = new JObject();
        jSeriesItem.Add(new JProperty("type", "area"));
        jSeriesItem.Add(new JProperty("name", "个人"));
        jSeriesItem.Add(new JProperty("color", "rgba(0,204,255,0.75)"));
        jSeriesItem.Add(new JProperty("data", jSeriesDataPerson));
        jSeries.Add(jSeriesItem);


        jChartInfo = new JObject();
        jChartInfo.Add(new JProperty("Category", jCategory));
        jChartInfo.Add(new JProperty("Series", jSeries));
        jItem.Add(new JProperty("KPC", jChartInfo));



        // 个人对年级
        jSeries = new JArray();
        jSeriesItem = new JObject();
        jSeriesDataPerson = new JArray();

        DataView dv;

        dv = ds.Tables[3].DefaultView;
        dv.Sort = "DDV ASC";
        DataTable dt = dv.ToTable();

        jCategory = new JArray();
        foreach (DataRow row in dt.Rows)     // ds.Tables(3).Rows
        {
            // jCategory.Add(row("KnowName"))
            jCategory.Add(row["RID"]);
            jSeriesDataPerson.Add(row["PDV"]);

            jSeriesDataGrade.Add(row["GDV"]);
        }

        // If jSeriesDataPerson.Count = 0 Then
        // jCategory.Add(0)
        // jSeriesDataPerson.Add("0")
        // jSeriesDataClass.Add("0")
        // End If

        jSeriesItem.Add(new JProperty("type", "line"));
        jSeriesItem.Add(new JProperty("name", "年级"));
        jSeriesItem.Add(new JProperty("color", "#ff0000"));
        jSeriesItem.Add(new JProperty("data", jSeriesDataGrade));
        jSeries.Add(jSeriesItem);

        jSeriesItem = new JObject();
        jSeriesItem.Add(new JProperty("type", "area"));
        jSeriesItem.Add(new JProperty("name", "个人"));
        jSeriesItem.Add(new JProperty("color", "rgba(0,204,255,0.75)"));
        jSeriesItem.Add(new JProperty("data", jSeriesDataPerson));
        jSeries.Add(jSeriesItem);


        jChartInfo = new JObject();
        jChartInfo.Add(new JProperty("Category", jCategory));
        jChartInfo.Add(new JProperty("Series", jSeries));

        jItem.Add(new JProperty("KPG", jChartInfo));

        // ----------------能力点
        jCategory = new JArray();
        jSeriesDataPerson = new JArray();
        jSeriesDataClass = new JArray();
        jSeriesDataGrade = new JArray();
        jSeries = new JArray();


        // For iCounter = 1 To ds.Tables(4).Columns.Count
        // jCategory.Add(iCounter.ToString())
        // Next

        iCounter = 1;
        KBInfo = "";
        KHInfo = "";

        foreach (DataRow row in ds.Tables[4].Rows)
        {
            // jCategory.Add(row("KnowName"))
            jCategory.Add(row["RID"]);
            jSeriesDataPerson.Add(row["PDV"]);
            jSeriesDataClass.Add(row["BDV"]);
            // jSeriesDataGrade.Add(row("GDV"))
            if (row["DDVF"] < 0)
                KBInfo = KBInfo + "<li><span style='color:#f00;font-weight:700;'>K" + iCounter.ToString() + "</span>：" + row["NLName"] + "</li>";
            if (row["DDVF"] > 0)
                KHInfo = KHInfo + "<li><span style='color:#f00;font-weight:700;'>K" + iCounter.ToString() + "</span>：" + row["NLName"] + "</li>";

            iCounter += 1;
        }

        // If jSeriesDataPerson.Count = 0 Then
        // jCategory.Add(0)
        // jSeriesDataPerson.Add("0")
        // jSeriesDataClass.Add("0")
        // End If

        jItem.Add(new JProperty("NLBInfo", KBInfo));
        jItem.Add(new JProperty("NLHInfo", KHInfo));

        jSeriesItem = new JObject();
        jSeriesItem.Add(new JProperty("type", "line"));
        jSeriesItem.Add(new JProperty("name", "班级"));
        jSeriesItem.Add(new JProperty("color", "#ff0000"));
        jSeriesItem.Add(new JProperty("data", jSeriesDataClass));
        jSeries.Add(jSeriesItem);

        jSeriesItem = new JObject();
        jSeriesItem.Add(new JProperty("type", "area"));
        jSeriesItem.Add(new JProperty("name", "个人"));
        jSeriesItem.Add(new JProperty("color", "rgba(0,204,255,0.75)"));
        jSeriesItem.Add(new JProperty("data", jSeriesDataPerson));
        jSeries.Add(jSeriesItem);


        jChartInfo = new JObject();
        jChartInfo.Add(new JProperty("Category", jCategory));
        jChartInfo.Add(new JProperty("Series", jSeries));
        jItem.Add(new JProperty("NLPC", jChartInfo));



        // 个人对年级
        jSeries = new JArray();
        jSeriesItem = new JObject();
        jSeriesDataPerson = new JArray();



        dv = ds.Tables[4].DefaultView;
        dv.Sort = "DDV ASC";
        dt = dv.ToTable();
        jCategory = new JArray();

        foreach (DataRow row in dt.Rows)     // ds.Tables(3).Rows
        {
            // jCategory.Add(row("KnowName"))
            jCategory.Add(row["RID"]);
            jSeriesDataPerson.Add(row["PDV"]);

            jSeriesDataGrade.Add(row["GDV"]);
        }

        // If jSeriesDataPerson.Count = 0 Then
        // jCategory.Add(0)
        // jSeriesDataPerson.Add("0")
        // jSeriesDataGrade.Add("0")
        // End If

        jSeriesItem.Add(new JProperty("type", "line"));
        jSeriesItem.Add(new JProperty("name", "年级"));
        jSeriesItem.Add(new JProperty("color", "#ff0000"));
        jSeriesItem.Add(new JProperty("data", jSeriesDataGrade));
        jSeries.Add(jSeriesItem);

        jSeriesItem = new JObject();
        jSeriesItem.Add(new JProperty("type", "area"));
        jSeriesItem.Add(new JProperty("name", "个人"));
        jSeriesItem.Add(new JProperty("color", "rgba(0,204,255,0.75)"));
        jSeriesItem.Add(new JProperty("data", jSeriesDataPerson));
        jSeries.Add(jSeriesItem);


        jChartInfo = new JObject();
        jChartInfo.Add(new JProperty("Category", jCategory));
        jChartInfo.Add(new JProperty("Series", jSeries));

        jItem.Add(new JProperty("NLPG", jChartInfo));

        // --------------------能力点

        jTableData = new JArray();

        KBInfo = "";
        KHInfo = "";

        foreach (DataRow dr in ds.Tables[5].Rows)
        {
            jRowData = new JObject();
            foreach (DataColumn col in ds.Tables[5].Columns)
                jRowData.Add(new JProperty(col.ColumnName, dr[col.ColumnName]));

            if (dr["PDV"] < dr["GDV"])
            {
                if (dr["GDV"] < 60)
                    KBInfo = KBInfo + "<ul><li>" + dr["KnowName"] + "：</li><li>" + dr["Item_Name"] + "</li></ul>";
                else
                    KHInfo = KHInfo + "<ul><li>" + dr["KnowName"] + "：</li><li>" + dr["Item_Name"] + "</li></ul>";
            }

            jTableData.Add(jRowData);
        }

        jItem.Add(new JProperty("QNL", KBInfo));
        jItem.Add(new JProperty("QXX", KHInfo));


        jItem.Add(new JProperty("QData", jTableData));


        NGenStuReportByTestID = jItem.ToString();
    }


    private string GenHtmlReport(HttpContext context)
    {
        string iReportID; // = "2"
        string sTestID; // = "4"
        string sLessonID; // = "语文"
        string iStatType; // = "1"
        string sStatUnitID; // = "鄂南高中"

        string sUserID = context.Session["LoginUserID"].ToString();

        // StatUnitID  TestID LessonID StatLevel ReportID

        if (context.Request.QueryString["resData"] != null)
        {
            string sResData = context.Request.QueryString["resData"].ToString();
            string[] sPara = sResData.Split("|");
            sStatUnitID = sPara[0];
            sTestID = sPara[1];
            sLessonID = sPara[2];
            iStatType = sPara[3];
            iReportID = sPara[4];
        }



        string[] sT;

        string sGID;

        string sReportDate;

        sT = DBHelp.GetString("Select cast(TestGradeID as nvarchar) +'/'+convert( nvarchar(10),TestDate,121)  From tblTestList Where TestID=" + sTestID).Split("/");
        sGID = sT[0];
        sReportDate = sT[1];


        // 0ShowID 1SchoolID 2ReportCommand 3GenProcess 4Config 5 StatType  
        string sReportConfig;
        sReportConfig = "SELECT   [ReportTile]+'/'+ [ReportDesp]+'/'+ [ReportCommand]+'/'+ cast([GenProcess] as nvarchar)+'/'+isnull([Config],'') " + "  FROM  [tblReport] Where ShowID=" + iReportID;

        sReportConfig = DBHelp.GetString(sReportConfig);

        sT = sReportConfig.Split("/");

        string sReportTitle = sT[0];
        string sReportDesp = sT[1];


        string sReportCommand = sT[2];
        int iGenProcess = System.Convert.ToInt32(sT[3]);

        sReportConfig = sT[4];

        // pr_ReportSingleZTTJ 0     |@UserID 1|@TestID 2|@LessonID 3|@StatType 4|@StatUnit 5
        string[] sCommand = sReportCommand.ToString().Split("|");

        IDataParameter[] parm = new IDataParameter[sCommand.Length - 2 + 1];

        parm[0] = DBHelp.CreateParameter(sCommand[1].Trim(), sUserID);
        parm[1] = DBHelp.CreateParameter(sCommand[2].Trim(), sTestID);

        if (Strings.StrComp(sCommand[3].Trim(), "@LessonID", CompareMethod.Text) == 0)
            parm[2] = DBHelp.CreateParameter(sCommand[3].Trim(), sLessonID);
        else
            parm[2] = DBHelp.CreateParameter(sCommand[3].Trim(), sGID);

        parm[3] = DBHelp.CreateParameter(sCommand[4].Trim(), iStatType);  // "StatType
        parm[4] = DBHelp.CreateParameter(sCommand[5].Trim(), sStatUnitID);  // " SchoolID

        // 这是一个自动配置的参数--------------
        if (parm.Length > 5)
        {
            string[] sDefault;
            for (int iCounter = 6; iCounter <= sCommand.Length - 1; iCounter++)
            {
                sDefault = sCommand[iCounter].Trim().Split("?");
                parm[iCounter - 1] = DBHelp.CreateParameter(sDefault[0].Trim(), sDefault[1]);
            }
        }

        // -----------自动配置的参数

        string sCommandProc;
        sCommandProc = sCommand[0].Trim();
        DataSet ds;

        try
        {
            ds = DBHelp.GetDataSet(sCommandProc, parm);
        }
        catch (Exception ex)
        {
            return;
        }
        finally
        {
        }

        ds.Dispose();

        GenHtmlReport = StiHtmlReport.GenReportByHtml(ds, 0);  // JsonConvert.SerializeObject(sHtml)
    }



    public string GenStuAllTestInfo(HttpContext context)
    {
        string sXHID = "";

        IDataParameter[] parm = new IDataParameter[1] { };

        int iUserType;
        int.TryParse(context.Session["UserType"], ref iUserType);

        if (iUserType == -1)
            sXHID = context.Session["LoginUserID"].ToString();
        else if (context.Request.QueryString["XHID"] != null)
            sXHID = context.Request.QueryString["XHID"].ToString();


        parm[0] = DBHelp.CreateParameter("@XHID", sXHID);

        DataSet ds = DBHelp.GetDataSet("pr_GenStuAllTestInfo", parm);
        ds.Tables[0].TableName = "NewScore";

        ds.Tables[1].TableName = "LCZMark";  // LID)
        ds.Tables[2].TableName = "LCZLevel";

        ds.Tables[3].TableName = "KnowLID";    // LID
        ds.Tables[4].TableName = "KnowInfo";

        ds.Tables[5].TableName = "ReportTest";   // TestID        
        ds.Tables[6].TableName = "EQInfo";   // TestID LID



        // 必须带上False 否则 不能启用此约束 因为不是所有的值都具有相应的父值
        // ds.Relations.Add("LCChild", ds.Tables("LCLIDList").Columns("LID"), _
        // ds.Tables("LCInfo").Columns("LID"), False)


        ds.Relations.Add("KnowChild", ds.Tables["KnowLID"].Columns["LID"], ds.Tables["KnowInfo"].Columns["LID"], false);



        ds.Relations.Add("EQChild", new DataColumn[] { ds.Tables["ReportTest"].Columns["TestID"], ds.Tables["ReportTest"].Columns["LID"] }, new DataColumn[] { ds.Tables["EQInfo"].Columns["TestID"], ds.Tables["EQInfo"].Columns["LID"] }, false);


        DataRow rowParent, rowChild;


        JObject jAllItem = new JObject();

        JObject jItem = new JObject();

        JArray jArr = new JArray();

        string sUserName = "";
        if (ds.Tables[7].Rows.Count > 0)
            sUserName = ds.Tables[7].Rows[0][0].ToString();

        if (iUserType == -1)
        {
            jAllItem.Add(new JProperty("ISME", 1));
            jAllItem.Add(new JProperty("XHID", sXHID));
            jAllItem.Add(new JProperty("StuName", sUserName));
        }
        else
        {
            jAllItem.Add(new JProperty("ISME", 0));
            jAllItem.Add(new JProperty("XHID", sXHID));
            jAllItem.Add(new JProperty("StuName", sUserName));
        }

        foreach (var rowParent in ds.Tables["NewScore"].Rows)
        {
            jItem = new JObject();

            jItem.Add(new JProperty("LessonID", rowParent["LID"]));
            jItem.Add(new JProperty("TotalScore", rowParent["TotalScore"]));
            jItem.Add(new JProperty("AVMark", rowParent["average_mark"]));
            jItem.Add(new JProperty("MultiLesson", rowParent["MultiLesson"]));

            jArr.Add(jItem);
        }
        jAllItem.Add(new JProperty("LastScore", jArr));
        // MsgBox(jArr.ToString())

        JObject jItemChild = new JObject();

        jItem = new JObject();
        string[] sColumnInfo;


        jArr = new JArray();
        int iCounter;



        jItem = new JObject();
        jItem.Add(new JProperty("LessonID", "Z分数"));

        ;/* Cannot convert RedimClauseSyntax, System.InvalidCastException: Unable to cast object of type 'Microsoft.CodeAnalysis.VisualBasic.Syntax.RangeArgumentSyntax' to type 'Microsoft.CodeAnalysis.VisualBasic.Syntax.SimpleArgumentSyntax'.
   at ICSharpCode.CodeConverter.CSharp.CommonConversions.<ConvertArrayBounds>b__20_0(ArgumentSyntax a)
   at System.Linq.Enumerable.WhereSelectEnumerableIterator`2.MoveNext()
   at System.Collections.Generic.List`1..ctor(IEnumerable`1 collection)
   at System.Linq.Enumerable.ToList[TSource](IEnumerable`1 source)
   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.MethodBodyVisitor.VisitRedimClause(RedimClauseSyntax node)
   at Microsoft.CodeAnalysis.VisualBasic.Syntax.RedimClauseSyntax.Accept[TResult](VisualBasicSyntaxVisitor`1 visitor)
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.Visit(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.ConvertWithTrivia(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.DefaultVisit(SyntaxNode node)

Input: 
sColumnInfo(0 To ds.Tables(1).Columns.Count - 1)

 */

        iCounter = 0;
        foreach (DataColumn col in ds.Tables[1].Columns)
        {
            sColumnInfo[iCounter] = col.ColumnName;
            iCounter += 1;
        }
        jItemChild = StiHtmlReport.GetHtmlChart(20, ds.Tables[1], sColumnInfo);

        // jItemChild = StiHtmlReport.GetHtmlChartForChildRows(20 + iCounter, rowParent.GetChildRows("LCChild", DataRowVersion.Original), sColumnInfo, 12)

        jItem.Add(new JProperty("Data", jItemChild));

        jArr.Add(jItem);


        // -----------ZLevel
        jItem = new JObject();
        jItem.Add(new JProperty("LessonID", "百分位"));

        ;/* Cannot convert RedimClauseSyntax, System.InvalidCastException: Unable to cast object of type 'Microsoft.CodeAnalysis.VisualBasic.Syntax.RangeArgumentSyntax' to type 'Microsoft.CodeAnalysis.VisualBasic.Syntax.SimpleArgumentSyntax'.
   at ICSharpCode.CodeConverter.CSharp.CommonConversions.<ConvertArrayBounds>b__20_0(ArgumentSyntax a)
   at System.Linq.Enumerable.WhereSelectEnumerableIterator`2.MoveNext()
   at System.Collections.Generic.List`1..ctor(IEnumerable`1 collection)
   at System.Linq.Enumerable.ToList[TSource](IEnumerable`1 source)
   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.MethodBodyVisitor.VisitRedimClause(RedimClauseSyntax node)
   at Microsoft.CodeAnalysis.VisualBasic.Syntax.RedimClauseSyntax.Accept[TResult](VisualBasicSyntaxVisitor`1 visitor)
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.Visit(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.ConvertWithTrivia(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.DefaultVisit(SyntaxNode node)

Input: 
sColumnInfo(0 To ds.Tables(2).Columns.Count - 1)

 */

        iCounter = 0;
        foreach (DataColumn col in ds.Tables[2].Columns)
        {
            sColumnInfo[iCounter] = col.ColumnName;
            iCounter += 1;
        }
        jItemChild = StiHtmlReport.GetHtmlChart(21, ds.Tables[2], sColumnInfo);

        // jItemChild = StiHtmlReport.GetHtmlChartForChildRows(20 + iCounter, rowParent.GetChildRows("LCChild", DataRowVersion.Original), sColumnInfo, 12)

        jItem.Add(new JProperty("Data", jItemChild));

        jArr.Add(jItem);
        // -------------


        jAllItem.Add(new JProperty("LCChart", jArr));
        // MsgBox(jArr.ToString())


        jArr = new JArray();
        string strTitle;

        foreach (var rowParent in ds.Tables["KnowLID"].Rows)
        {
            jItem = new JObject();
            strTitle = "";

            foreach (var rowChild in rowParent.GetChildRows("KnowChild", DataRowVersion.Original))
                strTitle = strTitle + "<li>" + rowChild["KnowName"].ToString() + "</li>";

            jItem.Add(new JProperty("LessonID", rowParent["LID"]));
            jItem.Add(new JProperty("KnowData", strTitle));

            jArr.Add(jItem);
        }

        jAllItem.Add(new JProperty("BadKnow", jArr));
        // MsgBox(jArr.ToString())

        // 考试报告 错题
        jArr = new JArray();



        JArray jArrChild = new JArray();

        foreach (var rowParent in ds.Tables["ReportTest"].Rows)
        {
            jItem = new JObject();

            jItem.Add(new JProperty("TestID", rowParent["TestID"].ToString()));
            jItem.Add(new JProperty("LessonID", rowParent["LID"].ToString()));

            strTitle = "<i class='icon-book'></i>" + rowParent["TestTitle"].ToString() + "_" + rowParent["TestDate"].ToString() + "(" + rowParent["LID"].ToString() + ")";

            jItem.Add(new JProperty("Title", strTitle));

            strTitle = "<ul class='list-group-item-meta'> <li>考试时间 <span>" + rowParent["TestDate"].ToString() + "</span> " + " 得分 <span class='a'>" + rowParent["TotalScore"].ToString() + "</span>| 年级平均成绩 <span class='a'>" + rowParent["Average_Mark"].ToString() + "</span> " + " |  <span class='d'></span> 失分</li> </ul>";

            jItem.Add(new JProperty("SubTitle", strTitle));


            jArrChild = new JArray();

            foreach (var rowChild in rowParent.GetChildRows("EQChild", DataRowVersion.Original))
            {
                jItemChild = new JObject();

                jItemChild.Add(new JProperty("QID", rowChild["ItemID"]));

                strTitle = rowChild["Item_Name"] + " <span>(" + rowChild["MaxScore"] + "分)</span> <span class='d'>【" + rowChild["Mark_Value"] - rowChild["MaxScore"] + "分】</span> ";

                jItemChild.Add(new JProperty("QTitle", strTitle));
                jItemChild.Add(new JProperty("KnowList", rowChild["KnowName"]));
                jItemChild.Add(new JProperty("NLList", rowChild["NLName"]));


                jArrChild.Add(jItemChild);
            }

            jItem.Add(new JProperty("EQInfo", jArrChild));

            jArr.Add(jItem);
        }

        jAllItem.Add(new JProperty("ReportList", jArr));
        // MsgBox(jArr.ToString())

        GenStuAllTestInfo = jAllItem.ToString();
    }

    public string GetQuesMedia(HttpContext context)
    {
        string sTestID;  // 'context.Session("LoginUserID").ToString()        
        string sLessonID;
        string sQuesID;


        if (context.Request.QueryString["TestID"] != null)
            sTestID = context.Request.QueryString["TestID"].ToString();

        if (context.Request.QueryString["LessonID"] != null)
            sLessonID = context.Request.QueryString["LessonID"].ToString();

        if (context.Request.QueryString["QuesID"] != null)
            sQuesID = context.Request.QueryString["QuesID"].ToString();


        IDataParameter[] parm = new IDataParameter[3] { };

        parm[0] = DBHelp.CreateParameter("@TestID", sTestID);
        parm[1] = DBHelp.CreateParameter("@LessonID", sLessonID);
        parm[2] = DBHelp.CreateParameter("@QuesID", sQuesID);

        DataSet ds = DBHelp.GetDataSet("pr_GetQuesMedia", parm);


        ds.Tables[0].TableName = "Cat";
        ds.Tables[1].TableName = "Media";
        ds.Tables[2].TableName = "KN";

        GetQuesMedia = JsonConvert.SerializeObject(ds);
    }


    public string GetClassTestLessonReport(HttpContext context)
    {
        string sTestID;
        string sLessonID;
        string sUnitID; // = "10麻城一中理17"



        if (context.Request.QueryString["TestID"] != null)
            sTestID = context.Request.QueryString["TestID"].ToString();

        if (context.Request.QueryString["LessonID"] != null)
            sLessonID = context.Request.QueryString["LessonID"].ToString();

        if (context.Request.QueryString["UnitID"] != null)
            sUnitID = context.Request.QueryString["UnitID"].ToString();

        IDataParameter[] parm = new IDataParameter[3] { };

        parm[0] = DBHelp.CreateParameter("@TestID", sTestID);
        parm[1] = DBHelp.CreateParameter("@LessonID", sLessonID);
        parm[2] = DBHelp.CreateParameter("@UnitID", sUnitID);


        DataSet ds = DBHelp.GetDataSet("pr_GenClassTestLessonReport", parm);
        ds.Tables[0].TableName = "Basic";
        ds.Tables[1].TableName = "FSD";
        ds.Tables[2].TableName = "Know";
        ds.Tables[3].TableName = "NLD";
        ds.Tables[4].TableName = "Ques";
        // 必须带上False 否则 不能启用此约束 因为不是所有的值都具有相应的父值
        ds.Relations.Add("ErrorList", ds.Tables[5].Columns["ItemID"], ds.Tables[6].Columns["ItemID"], false);

        DataRow rowParent, rowChild;

        JObject jItem = new JObject();

        JArray jChildData = new JArray();
        JArray jTableData = new JArray();
        JObject jRowData = new JObject();

        int iCounter;
        for (iCounter = 0; iCounter <= 4; iCounter++)
        {
            jTableData = new JArray();

            if (iCounter == 1)
            {
                jChildData = new JArray();
                jRowData = new JObject();
                foreach (var rowChild in ds.Tables[iCounter].Rows)
                {
                    jChildData.Add(rowChild[0].ToString());
                    jTableData.Add(rowChild[1]);
                }

                jRowData.Add(new JProperty("Category", jChildData));
                jRowData.Add(new JProperty("Data", jTableData));

                jItem.Add(new JProperty(ds.Tables[iCounter].TableName, jRowData));
            }
            else
            {
                foreach (var rowChild in ds.Tables[iCounter].Rows)
                {
                    jRowData = new JObject();

                    foreach (DataColumn col in ds.Tables[iCounter].Columns)
                        jRowData.Add(new JProperty(col.ColumnName, rowChild[col.ColumnName]));

                    jTableData.Add(jRowData);
                }

                jItem.Add(new JProperty(ds.Tables[iCounter].TableName, jTableData));
            }
        }

        jTableData = new JArray();


        JObject jRowDataParent;

        foreach (var rowParent in ds.Tables[5].Rows)
        {
            jRowDataParent = new JObject();

            // QName,  QTMark,A.ZeroCount,A.RightCount,A.CDif,C.GDif
            jRowDataParent.Add(new JProperty("QName", rowParent["QName"]));
            jRowDataParent.Add(new JProperty("QTMark", rowParent["QTMark"]));
            jRowDataParent.Add(new JProperty("ZeroCount", rowParent["ZeroCount"]));
            jRowDataParent.Add(new JProperty("RightCount", rowParent["RightCount"]));
            jRowDataParent.Add(new JProperty("CDif", rowParent["CDif"]));
            jRowDataParent.Add(new JProperty("GDif", rowParent["GDif"]));
            jRowDataParent.Add(new JProperty("ZSD", rowParent["ZSD"]));
            jRowDataParent.Add(new JProperty("NLD", rowParent["NLD"]));


            jChildData = new JArray();

            foreach (var rowChild in rowParent.GetChildRows("ErrorList", DataRowVersion.Original))
            {
                jRowData = new JObject();

                jRowData.Add(new JProperty("ItemStr", rowChild["ItemStr"]));
                jRowData.Add(new JProperty("EC", rowChild["EC"]));
                jRowData.Add(new JProperty("NameList", rowChild["NameList"]));

                jChildData.Add(jRowData);
            }

            jRowDataParent.Add(new JProperty("ErrorList", jChildData));

            jTableData.Add(jRowDataParent);
        }

        jItem.Add(new JProperty("QInfo", jTableData));


        GetClassTestLessonReport = jItem.ToString();
    }


    public string ChangUserPW(HttpContext context)
    {
        string sUserID = context.Session["LoginUserID"].ToString();
        string sOldPW = "";
        string sNewPW;

        if (context.Request.QueryString["OldPWD"] != null)
            sOldPW = context.Request.QueryString["OldPWD"].ToString();

        if (context.Request.QueryString["NewPWD"] != null)
            sNewPW = context.Request.QueryString["NewPWD"].ToString();

        string strSQL;
        int iUserType;
        int.TryParse(context.Session["UserType"].ToString(), ref iUserType);

        if (iUserType == -1)
            strSQL = "SET NOCOUNT ON Update tblStudentInfo Set UPW='" + sNewPW + "' Where XHID='" + sUserID + "' And UPW='" + sOldPW + "'  SELECT @@ROWCOUNT SET NOCOUNT OFF";
        else
            strSQL = "SET NOCOUNT ON Update tblUserManage Set UserPW='" + sNewPW + "' Where UserID='" + sUserID + "' And UserPW='" + sOldPW + "'  SELECT @@ROWCOUNT SET NOCOUNT OFF";

        iUserType = DBHelp.GetScaler(strSQL);


        return iUserType.ToString();
    }

    public void DownLoadFiles(HttpContext context)
    {
        string sIDList = "";

        if (context.Request.QueryString["IDList"] != null)
            sIDList = context.Request.QueryString["IDList"].ToString();

        using (ZipFile zip = new ZipFile())
        {
            zip.AlternateEncodingUsage = ZipOption.AsNecessary;


            string strSQL;

            strSQL = "Select MediaURL from dbo.SplitStr('" + sIDList + "',',') A inner join tblTestMedia b ON A.[Value]=B.SourceID";
            DataSet ds = DBHelp.GetDataSet(strSQL);

            string sFile;


            foreach (DataRow row in ds.Tables[0].Rows)
            {
                sFile = context.Server.MapPath(row[0]);
                zip.AddFile(sFile, "");
            }

            context.Response.Buffer = true;

            context.Response.Clear();
            // context.Response.BufferOutput = False
            string zipName = HttpUtility.UrlEncode(String.Format("Zip_{0}.zip", DateTime.Now.ToString("yyyy-MMM-dd-HHmmss")), System.Text.Encoding.UTF8);
            context.Response.ContentType = "application/zip";
            context.Response.AddHeader("content-disposition", "attachment; filename=" + zipName);
            zip.Save(context.Response.OutputStream);
            context.Response.Flush();
            context.Response.End();
        }
    }

    public string GetStuPaperURL(HttpContext context)
    {
        string sTestID;   // context.Session("LoginUserID").ToString()
        string sXHID;
        string sLessonID;



        if (context.Request.QueryString["TestID"] != null)
            sTestID = context.Request.QueryString["TestID"].ToString();

        int iUserType;
        int.TryParse(context.Session["UserType"], ref iUserType);

        if (iUserType == -1)
            sXHID = context.Session["LoginUserID"].ToString();
        else if (context.Request.QueryString["XHID"] != null)
            sXHID = context.Request.QueryString["XHID"].ToString();

        if (context.Request.QueryString["LessonID"] != null)
            sLessonID = context.Request.QueryString["LessonID"].ToString();


        IDataParameter[] parm = new IDataParameter[3] { };

        parm[0] = DBHelp.CreateParameter("@TestID", sTestID);
        parm[1] = DBHelp.CreateParameter("@LessonID", sLessonID);
        parm[2] = DBHelp.CreateParameter("@XHID", sXHID);


        DataSet ds = DBHelp.GetDataSet("pr_GetStuPaperURL", parm);
        return JsonConvert.SerializeObject(ds.Tables[0]);
    }

    public string GetClassLessonTestList(HttpContext context)
    {
        string sUserID = context.Session["LoginUserID"].ToString();
        string sTestID;   // context.Session("LoginUserID").ToString()
        string sStatUnitID;
        string sLessonID;



        if (context.Request.QueryString["TestID"] != null)
            sTestID = context.Request.QueryString["TestID"].ToString();



        if (context.Request.QueryString["StatUnitID"] != null)
            sStatUnitID = context.Request.QueryString["StatUnitID"].ToString();


        if (context.Request.QueryString["LessonID"] != null)
            sLessonID = context.Request.QueryString["LessonID"].ToString();


        IDataParameter[] parm = new IDataParameter[4] { };
        parm[0] = DBHelp.CreateParameter("@UserID", sUserID);
        parm[1] = DBHelp.CreateParameter("@TestID", sTestID);
        parm[2] = DBHelp.CreateParameter("@LessonID", sLessonID);
        parm[3] = DBHelp.CreateParameter("@ClassID", sStatUnitID);


        DataSet ds = DBHelp.GetDataSet("pr_GetClassLessonTestList", parm);
        ds.Tables[0].TableName = "Basic";
        ds.Tables[1].TableName = "Know";
        ds.Tables[2].TableName = "ClassList";
        return JsonConvert.SerializeObject(ds);
    }


    public string LogOffSys(HttpContext context)
    {
        context.Session["LoginUserID"] = null;
        context.Session["UserType"] = null;
        string sURL = "Default.aspx";
        return JsonConvert.SerializeObject(sURL);
    }


    private string GenHtmlReportByPage(HttpContext context)
    {
        string iReportID; // = "2"
        string sTestID; // = "4"
        string sLessonID; // = "语文"
        string iStatType; // = "1"
        string sStatUnitID; // = "鄂南高中"
        string sFK = "-1";

        string sUserID = context.Session["LoginUserID"].ToString();

        // StatUnitID  TestID LessonID StatLevel ReportID
        if (context.Request.QueryString["WL"] != null)
            sFK = context.Server.UrlDecode(context.Request.QueryString["WL"].ToString());
        if (context.Request.QueryString["resData"] != null)
        {
            string sResData = context.Server.UrlDecode(context.Request.QueryString["resData"].ToString());
            string[] sPara = sResData.Split("|");
            sStatUnitID = sPara[0];
            sTestID = sPara[1];
            sLessonID = sPara[2];
            iStatType = sPara[3];
            iReportID = sPara[4];
        }

        int echo = int.Parse(context.Request.Params["sEcho"]);
        int displayLength = int.Parse(context.Request.Params["iDisplayLength"]);
        int displayStart = int.Parse(context.Request.Params["iDisplayStart"]);

        displayStart += 1;

        string[] sT;


        // 0ShowID 1SchoolID 2ReportCommand 3GenProcess 4Config 5 StatType  
        string sReportConfig;
        sReportConfig = "SELECT   [ReportTile]+'/'+ [ReportDesp]+'/'+ [ReportCommand]+'/'+ cast([GenProcess] as nvarchar)+'/'+isnull([Config],'') " + "  FROM  [tblReport] Where ShowID=" + iReportID;

        sReportConfig = DBHelp.GetString(sReportConfig);

        sT = sReportConfig.Split("/");

        string sReportTitle = sT[0];
        string sReportDesp = sT[1];


        string sReportCommand = sT[2];
        int iGenProcess = System.Convert.ToInt32(sT[3]);

        sReportConfig = sT[4];

        // pr_ReportSingleZTTJ 0     |@UserID 1|@TestID 2|@LessonID 3|@StatType 4|@StatUnit 5
        string[] sCommand = sReportCommand.ToString().Split("|");

        IDataParameter[] parm = new IDataParameter[sCommand.Length - 2 + 1];

        parm[0] = DBHelp.CreateParameter(sCommand[1].Trim(), sUserID);
        parm[1] = DBHelp.CreateParameter(sCommand[2].Trim(), sTestID);

        if (Strings.StrComp(sCommand[3].Trim(), "@LessonID", CompareMethod.Text) == 0)
            parm[2] = DBHelp.CreateParameter(sCommand[3].Trim(), sLessonID);
        else
            parm[2] = DBHelp.CreateParameter(sCommand[3].Trim(), "");

        parm[3] = DBHelp.CreateParameter(sCommand[4].Trim(), iStatType);  // "StatType
        parm[4] = DBHelp.CreateParameter(sCommand[5].Trim(), sStatUnitID);  // " SchoolID

        // 这是一个自动配置的参数--------------
        if (parm.Length > 5)
        {
            string[] sDefault;
            for (int iCounter = 6; iCounter <= sCommand.Length - 1; iCounter++)
            {
                sDefault = sCommand[iCounter].Trim().Split("?");
                switch (sDefault[0].Trim())
                {
                    case "@DisplayStart":
                        {
                            parm[iCounter - 1] = DBHelp.CreateParameter(sDefault[0].Trim(), displayStart);
                            break;
                        }

                    case "@DisplayLen":
                        {
                            parm[iCounter - 1] = DBHelp.CreateParameter(sDefault[0].Trim(), displayLength);
                            break;
                        }

                    case "@FK":
                        {
                            parm[iCounter - 1] = DBHelp.CreateParameter(sDefault[0].Trim(), sFK);
                            break;
                        }

                    default:
                        {
                            parm[iCounter - 1] = DBHelp.CreateParameter(sDefault[0].Trim(), sDefault[1]);
                            break;
                        }
                }
            }
        }

        // -----------自动配置的参数

        string sCommandProc;
        sCommandProc = sCommand[0].Trim();
        DataSet ds;

        try
        {
            ds = DBHelp.GetDataSet(sCommandProc, parm);
        }
        catch (Exception ex)
        {
            return;
        }
        finally
        {
        }

        ds.Dispose();

        GenHtmlReportByPage = StiHtmlReport.GetServerSideData(ds.Tables[0], echo);
    }


    private void ExportReport(HttpContext context)
    {
        string iReportID; // = "2"
        string sTestID; // = "4"
        string sLessonID; // = "语文"
        string iStatType; // = "1"
        string sStatUnitID; // = "鄂南高中"

        string iFileType;

        string sUserID = context.Session["LoginUserID"].ToString();

        // StatUnitID  TestID LessonID StatLevel ReportID

        if (context.Request.QueryString["resData"] != null)
        {
            string sResData = context.Request.QueryString["resData"].ToString();
            string[] sPara = sResData.Split("|");
            sStatUnitID = sPara[0];
            sTestID = sPara[1];
            sLessonID = sPara[2];
            iStatType = sPara[3];
            iReportID = sPara[4];
            iFileType = sPara[5];
        }



        string[] sT;

        string sGID;

        string sReportDate;

        sT = DBHelp.GetString("Select cast(TestGradeID as nvarchar) +'/'+convert( nvarchar(10),TestDate,121)  From tblTestList Where TestID=" + sTestID).Split("/");
        sGID = sT[0];
        sReportDate = sT[1];


        // 0ShowID 1SchoolID 2ReportCommand 3GenProcess 4Config 5 StatType  
        string sReportConfig;
        sReportConfig = "SELECT   [ReportTile]+'/'+ [ReportDesp]+'/'+ [ReportCommand]+'/'+ cast([GenProcess] as nvarchar)+'/'+isnull([Config],'') " + "  FROM  [tblReport] Where ShowID=" + iReportID;

        sReportConfig = DBHelp.GetString(sReportConfig);

        sT = sReportConfig.Split("/");

        string sReportTitle = sT[0];
        string sReportDesp = sT[1];


        string sReportCommand = sT[2];
        int iGenProcess = System.Convert.ToInt32(sT[3]);

        sReportConfig = sT[4];

        // pr_ReportSingleZTTJ 0     |@UserID 1|@TestID 2|@LessonID 3|@StatType 4|@StatUnit 5
        string[] sCommand = sReportCommand.ToString().Split("|");

        IDataParameter[] parm = new IDataParameter[sCommand.Length - 2 + 1];

        parm[0] = DBHelp.CreateParameter(sCommand[1].Trim(), sUserID);
        parm[1] = DBHelp.CreateParameter(sCommand[2].Trim(), sTestID);

        if (Strings.StrComp(sCommand[3].Trim(), "@LessonID", CompareMethod.Text) == 0)
            parm[2] = DBHelp.CreateParameter(sCommand[3].Trim(), sLessonID);
        else
            parm[2] = DBHelp.CreateParameter(sCommand[3].Trim(), sGID);

        parm[3] = DBHelp.CreateParameter(sCommand[4].Trim(), iStatType);  // "StatType
        parm[4] = DBHelp.CreateParameter(sCommand[5].Trim(), sStatUnitID);  // " SchoolID

        // 这是一个自动配置的参数--------------
        if (parm.Length > 5)
        {
            string[] sDefault;
            for (int iCounter = 6; iCounter <= sCommand.Length - 1; iCounter++)
            {
                sDefault = sCommand[iCounter].Trim().Split("?");
                parm[iCounter - 1] = DBHelp.CreateParameter(sDefault[0].Trim(), sDefault[1]);
            }
        }

        // -----------自动配置的参数

        string sCommandProc;
        sCommandProc = sCommand[0].Trim();
        DataSet ds;

        try
        {
            ds = DBHelp.GetDataSet(sCommandProc, parm);
        }
        catch (Exception ex)
        {
            return;
        }
        finally
        {
        }

        ds.Dispose();

        string sConfig = sReportConfig.ToString();
        // Dim r As Regex
        // r = New Regex("(W\[\d{3}\])")
        int iPaperConfig = 100;
        string sPattern = @"W\[(\d{3})\]";

        System.Text.RegularExpressions.Match mStr = Regex.Match(sConfig, sPattern);

        if (mStr.Groups.Count > 0)
            iPaperConfig = int.Parse(mStr.Groups[1].Value);

        // For Each Match As Match In r.Matches(sConfig)
        // iPaperConfig = Match.Value.Substring(2, 3)
        // Next

        StiMReport.GenReportByTable(context, ds, sReportTitle, sReportDesp, iPaperConfig, iFileType);
    }
}
