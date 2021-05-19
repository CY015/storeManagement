using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace 订货系统
{
    class Supply
    {
        //更新供应商信息
        public int updateInfo(string sID, string sName, string sCon)
        {
            string upCondition;
            ConnSQL upCon = new ConnSQL();
            try
            {
                if (sName != "")
                {
                    upCondition = "UPDATE Supply SET cy_sName='" + sName + "' WHERE cy_sID='" + sID + "'";
                    upCon.ExecuteUpdate(upCondition);
                }
                if (sCon != "")
                {
                    upCondition = "UPDATE Supply SET cy_sContact='" + sCon + "' WHERE cy_sID='" + sID + "'";
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

        //插入新供应商信息
        public int insertInfo(string sID, string sName, string sCon)
        {
            string inCondition;
            ConnSQL inCon = new ConnSQL();

            inCondition = "INSERT INTO Supply(cy_sID,cy_sName,cy_sContact) VALUES('" + sID + "','" + sName + "','" + sCon + "')";
            inCon.ExecuteUpdate(inCondition);

            return 1;
        }
    }
}
