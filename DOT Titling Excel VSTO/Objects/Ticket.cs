using System;

namespace DOT_Titling_Excel_VSTO
{
    class Ticket : IComparable
    {
        #region Private Members
        private string _ID;
        private string _Type;
        private string _Summary;
        private string _Status;
        private string _Sprint;

        #endregion

        #region Properties
        public string ID
        {
            get { return _ID; }
            set { _ID = value; }
        }
        public string Type
        {
            get { return _Type; }
            set { _Type = value; }
        }
        public string Summary
        {
            get { return _Summary; }
            set { _Summary = value; }
        }
        public string Status
        {
            get { return _Status; }
            set { _Status = value; }
        }
        public string Sprint
        {
            get { return _Sprint; }
            set { _Sprint = value; }
        }
        #endregion


        #region Contructors
        public Ticket(string id, string type, string summary, string status, string sprint)
        {
            _ID = id;
            _Type = type;
            _Summary = summary;
            _Status = status;
            _Sprint = sprint;
        }
        #endregion


        #region IComparable Members
        public int CompareTo(object obj)
        {
            if (obj is Ticket)
            {
                Ticket t = (Ticket)obj;
                return _ID.CompareTo(t.ID);
            }
            else
                throw new ArgumentException("Object is not a Ticket.");
        }
        #endregion
    }
}

