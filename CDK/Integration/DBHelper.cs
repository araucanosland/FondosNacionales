using CDK.Data;

namespace CDK.Integration
{
    public class DBHelper
    {
        #region Instancias Privadas

        private static DBHelperBase _instanceInterna;
        

        #endregion

        #region Instancias Singleton

        public static DBHelperBase InstanceInterna
        {
            get { return (_instanceInterna = _instanceInterna ?? new DBHelperBase("CN_INTERNA")); }
        }


        
        #endregion
    }
}
