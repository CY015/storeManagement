using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;

namespace 订货系统
{
    /// <summary>
    /// 零件类
    /// </summary>
    /// 
    class Component
    {
        //更新零件信息
        public int updateInfo(string cID, string cName,string cType, float cPrice, int cLess, string csID)
        {
            string upCondition;
            ConnSQL upCon = new ConnSQL();
       
            try
            {
                if (cName != "")
                {
                    upCondition = "UPDATE Component SET cy_cName='" + cName + "' WHERE cy_cID='" + cID + "'";
                    upCon.ExecuteUpdate(upCondition);
                }
                if (cType != "")
                {
                    upCondition = "UPDATE Component SET cy_cType='" + cType + "' WHERE cy_cID='" + cID + "'";
                    upCon.ExecuteUpdate(upCondition);
                }
                //if (cPrice != -0.1F || cPrice != 0)
                if (cPrice>0)
                {
                    upCondition = "UPDATE Component SET cy_cPrice='" + cPrice + "' WHERE cy_cID='" + cID + "'";
                    upCon.ExecuteUpdate(upCondition);
                }
                if (cLess != -1)
                {
                    upCondition = "UPDATE Component SET cy_cMinNum='" + cLess + "' WHERE cy_cID='" + cID + "'";
                    upCon.ExecuteUpdate(upCondition);
                }
                if (csID != "")
                {
                    upCondition = "UPDATE Component SET cy_sID='" + csID + "' WHERE cy_cID='" + cID + "'";
                    upCon.ExecuteUpdate(upCondition);
                }
            }
            catch (Exception)
            {
                return 0;
                throw;
            }
            return 1;
        }

        //插入零件信息
        public int insertInfo(string cID, string cName, string cType, float cPrice, int cLess, string csID)
        {
            string inCondition;
            ConnSQL inCon = new ConnSQL();

            inCondition = "INSERT INTO Component(cy_cID,cy_cName,cy_cType,cy_cPrice,cy_cMinNum,cy_sID,cy_cNum) VALUES('" + cID + "','" + cName + "','" + cType + "','" + cPrice + "','" + cLess + "','" + csID + "',0)";
            inCon.ExecuteUpdate(inCondition);

            return 1;
        }
    }
}
