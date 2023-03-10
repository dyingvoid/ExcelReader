namespace ExcelReader;

public static class Tests
{
    public static List<Tuple<object?, object?>> PublicInstancePropertiesNotEqual<T>(T self, T to, params string[] ignore) 
        where T : class
    {
        var list = new List<Tuple<object?, object?>>();
        
        if (self != null && to != null)
        {
            Type type = typeof(T);
            List<string> ignoreList = new List<string>(ignore);
            foreach (System.Reflection.PropertyInfo pi in 
                     type.GetProperties(System.Reflection.BindingFlags.Public | 
                                        System.Reflection.BindingFlags.Instance))
            {
                if (!ignoreList.Contains(pi.Name))
                {
                    object selfValue = type.GetProperty(pi.Name).GetValue(self, null);
                    object toValue = type.GetProperty(pi.Name).GetValue(to, null);

                    if (selfValue != toValue && (selfValue == null || !selfValue.Equals(toValue)))
                    {
                        list.Add(Tuple.Create(selfValue, toValue));
                    }
                }
            }
        }

        return list;
    }
}