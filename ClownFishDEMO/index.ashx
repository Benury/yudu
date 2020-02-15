<%@ WebHandler Language="C#" Class="index" %>

using System;
using System.Web;
using System.Text;
using System.Collections.Generic;
using System.Linq;

using System.Reflection;
using ClownFish;
public class index : IHttpHandler {

    public void ProcessRequest (HttpContext context) {
        string jsonResult = "";
        string requestType = context.Request.Params["HTTP_X_REQUESTED_WITH"];
        var method = context.Request["method"] ?? "";
        if (string.IsNullOrEmpty(method)) {
            return;
        }
        Type t = typeof(index);
        MethodInfo mt = t.GetMethod(method);
        if (mt == null) {
            return;
        }else
        {
            object res = (object)mt.Invoke(null,new HttpContext[] { context } );
            jsonResult = Newtonsoft.Json.JsonConvert.SerializeObject(res);
            context.Response.ContentType = "appliction/json";
            context.Response.Write(jsonResult);
            context.Response.End();
        }

    }

    public static object getUsers(HttpContext context)
    {
        using( DbContext sqlite = new DbContext("sqlite") ) {

            List<Users> users  = DbHelper.FillList<Users>("GetUsers", new {

            }, sqlite);
            return users;
        }
    }

    public static object insertUser(HttpContext context) {


        using( DbContext sqlite = new DbContext("sqlite") ) {
            Users user = new Users
            {
                xiang = context.Request.Form["xiang"],
                cun = context.Request.Form["cun"],
                sfzh = context.Request.Form["sfzh"],
                phone =  context.Request.Form["phone"].ToString(),
                edate = Convert.ToDateTime( context.Request.Form["edate"]),
                back_address = context.Request.Form["back_address"],
                back_type = context.Request.Form["back_type"],
                helthy = context.Request.Form["helthy"],
                mname = context.Request.Form["mname"],
                mjob = context.Request.Form["mjob"],
                mphone = context.Request.Form["mphone"].ToString(),
                desc = context.Request.Form["desc"],
                status = 0
            };
            string query = string.Format("select * from users where sfzh = {0}", context.Request.Form["sfzh"]);

            Users u = DbHelper.GetDataItem<Users>(query, null, sqlite, CommandKind.SqlTextNoParams);
            if (u == null)
            {
                int a = DbHelper.ExecuteScalar<int>("InsertUser", user, sqlite);
                return new {
                    data = new
                    {
                        status = 1
                    }
                };
            }
            else {
                return new {
                    data = new
                    {
                        status = 0
                    }
                };
            }


        }


    }
    public bool IsReusable {
        get {
            return false;
        }
    }

}