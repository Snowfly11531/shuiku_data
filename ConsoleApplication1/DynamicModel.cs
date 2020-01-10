using System;
using System.Collections.Generic;
using System.Dynamic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApplication1
{
    class DynamicModel:DynamicObject
    {
        private string propertyName;
        public string PropertyName
        {
            get { return propertyName; }
            set { propertyName = value; }
        }
        Dictionary<string, object> dicProperty
            = new Dictionary<string, object>();
        public Dictionary<string, object> DicProperty
        {
            get
            {
                return dicProperty;
            }
        }
        public int Count
        {
            get
            {
                return dicProperty.Count;
            }
        }
        public List<String> GetProperty()
        {
            return DicProperty.Select(i => i.Key).ToList();
        }

        public List<Object> GetValue()
        {
            return DicProperty.Select(i => i.Value).ToList();
        }

        public override bool TryGetMember(
            GetMemberBinder binder, out object result)
        {
            string name = binder.Name;
            return dicProperty.TryGetValue(name, out result);
        }
        public override bool TrySetMember(
            SetMemberBinder binder, object value)
        {
            if (binder.Name == "Property")
            {
                dicProperty[PropertyName] = value;
            }
            else
            {
                dicProperty[binder.Name] = value;
            }
            return true;
        }
    }
}
