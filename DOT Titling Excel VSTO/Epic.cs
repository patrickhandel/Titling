using System;

namespace DOT_Titling_Excel_VSTO
{
    class Epic : IComparable
    {
        #region Private Members
        private string _EpicName;
        private string _ReleaseName;
        private Int32 _ReleaseNumber;
        private Int32 _Priority;
        private string _Status;
        #endregion

        #region Properties
        public string EpicName
        {
            get { return _EpicName; }
            set { _EpicName = value; }
        }
        public string ReleaseName
        {
            get { return _ReleaseName; }
            set { _ReleaseName = value; }
        }
        public Int32 ReleaseNumber
        {
            get { return _ReleaseNumber; }
            set { _ReleaseNumber = value; }
        }

        public Int32 Priority
        {
            get { return _Priority; }
            set { _Priority = value; }
        }

        public string Status
        {
            get { return _Status; }
            set { _Status = value; }
        }
        #endregion


        #region Contructors
        public Epic(string epicName, string releaseName, Int32 releaseNumber, Int32 priority, string status)
        {
            _EpicName = epicName;
            _ReleaseName = releaseName;
            _ReleaseNumber = releaseNumber;
            _Priority = priority;
            _Status = status;
        }
        #endregion

        //public override string ToString()
        //{

        //    return String.Format(“{ 0}{ 1}, Age = { 2}“, _firstname,

        //    _lastname, _age.ToString());

        //}

        #region IComparable Members
        public int CompareTo(object obj)
        {
            if (obj is Epic)
            {
                Epic e = (Epic)obj;
                return _Priority.CompareTo(e.Priority);
            }
            else
                throw new ArgumentException("Object is not an Epic.");
        }
        #endregion
    }
}
