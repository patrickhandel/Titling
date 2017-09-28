using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DOT_Titling_Excel_VSTO
{
    class Release : IComparable
    {
        #region Private Members
        private Int32 _Number;
        private string _Name;
        private Int32 _DevSprintFrom;
        private Int32 _DevSprintTo;
        private Int32 _UATSprintFrom;
        private Int32 _UATSprintTo;
        private string _Status;
        #endregion

        #region Properties
        public Int32 Number
        {
            get { return _Number; }
            set { _Number = value; }
        }
        public string Name
        {
            get { return _Name; }
            set { _Name = value; }
        }
        public Int32 DevSprintFrom
        {
            get { return _DevSprintFrom; }
            set { _DevSprintFrom = value; }
        }
        public Int32 DevSprintTo
        {
            get { return _DevSprintTo; }
            set { _DevSprintTo = value; }
        }
        public Int32 UATSprintFrom
        {
            get { return _UATSprintFrom; }
            set { _UATSprintFrom = value; }
        }
        public Int32 UATSprintTo
        {
            get { return _UATSprintTo; }
            set { _UATSprintTo = value; }
        }
        public string Status
        {
            get { return _Status; }
            set { _Status = value; }
        }
        #endregion


        #region Contructors
        public Release(Int32 number, string name, Int32 devSprintFrom, Int32 devSprintTo, Int32 uatSprintFrom, Int32 uatSprintTo, string status)
        {
            _Number = number;
            _Name = name;
            _DevSprintFrom = devSprintFrom;
            _DevSprintTo = devSprintTo;
            _UATSprintFrom = uatSprintFrom;
            _UATSprintTo = uatSprintTo;
            _Status = status;
        }
        #endregion


        #region IComparable Members
        public int CompareTo(object obj)
        {
            if (obj is Release)
            {
                Release r = (Release)obj;
                return _Number.CompareTo(r.Number);
            }
            else
                throw new ArgumentException("Object is not a Release.");
        }
        #endregion
    }
}

