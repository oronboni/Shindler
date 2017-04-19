using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SchindlerManageReports
{
    class Information
    {
        string checktype;

        int count;

        public int Count
        {
            get
            {
                return count;
            }

            set
            {
                count = value;
            }
        }

        public string Checktype
        {
            get
            {
                return checktype;
            }

            set
            {
                checktype = value;
            }
        }

        public Information(string checktype)
        {
            Count = 1;
            this.Checktype = checktype;
        }
    }

    class TempInformation
    {
        string checktype;
        string checknum;

        public string Checktype
        {
            get
            {
                return checktype;
            }

            set
            {
                checktype = value;
            }
        }

        public string Checknum
        {
            get
            {
                return checknum;
            }

            set
            {
                checknum = value;
            }
        }

        public TempInformation(string checknum,string checktype)
        {
            this.checknum = checknum;
            this.Checktype = checktype;
        }
    }
}
