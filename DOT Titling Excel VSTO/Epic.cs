using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DOT_Titling_Excel_VSTO
{
    class Epic : IComparable
    {
        #region Private Members
        private string _EpicName;
        private string _ReleaseName;
        private Int32 _Release;
        private Int32 _SprintFrom;
        private Int32 _SprintTo;
        private Int32 _Priority;
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
        public Int32 Release
        {
            get { return _Release; }
            set { _Release = value; }
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
        public Int32 Priority
        {
            get { return _Priority; }
            set { _Priority = value; }
        }
        #endregion


        #region Contructors
        public Epic(string epicName, string releaseName, Int32 release, Int32 sprintFrom, Int32 sprintTo, Int32 priority)
        {
            _EpicName = epicName;
            _ReleaseName = releaseName;
            _Release = release;
            _SprintFrom = sprintFrom;
            _SprintTo = sprintTo;
            _Priority = priority;
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
