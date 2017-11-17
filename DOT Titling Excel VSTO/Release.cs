using System;

namespace DOT_Titling_Excel_VSTO
{
    class Release : IComparable
    {
        #region Private Members
        private Int32 _Number;
        private string _Name;
        private string _Status;
        private string _MidLong;
        private Int32 _SprintFrom;
        private Int32 _SprintTo;
        private Int32 _UATSprintFrom;
        private Int32 _UATSprintTo;
        private Int32 _VendorSprint;

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
        public string MidLong
        {
            get { return _MidLong; }
            set { _MidLong = value; }
        }
        public Int32 SprintFrom
        {
            get { return _SprintFrom; }
            set { _SprintFrom = value; }
        }
        public Int32 SprintTo
        {
            get { return _SprintTo; }
            set { _SprintTo = value; }
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
        public Int32 VendorSprint
        {
            get { return _VendorSprint; }
            set { _VendorSprint = value; }
        }

        public string Status
        {
            get { return _Status; }
            set { _Status = value; }
        }
        #endregion


        #region Contructors
        public Release(Int32 number, string name, string midLong, Int32 sprintFrom, Int32 sprintTo, Int32 uatSprintFrom, Int32 uatSprintTo, Int32 vendorSprint, string status)
        {
            _Number = number;
            _Name = name;
            _MidLong = midLong;
            _SprintFrom = sprintFrom;
            _SprintTo = sprintTo;
            _UATSprintFrom = uatSprintFrom;
            _UATSprintTo = uatSprintTo;
            _VendorSprint = vendorSprint;
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

