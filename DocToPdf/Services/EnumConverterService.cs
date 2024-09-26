using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocToPdf.Services
{
    public class EnumConverterService
    {/// <summary>
     /// To find Enum Value by enum index.
     /// </summary>
     /// <typeparam name="TContext">Enum Type</typeparam>
     /// <param name="index">Enum item index</param>
     /// <returns></returns>

        public static TContext FindEnumValue<TContext>(int index)
        {
            return (TContext)Enum.ToObject(typeof(TContext), index);
        }

        /// <summary>
        /// To find Enum Value by string.
        /// </summary>
        /// <typeparam name="TContext">Enum Type</typeparam>
        /// <param name="str">Enum item string</param>
        /// <returns></returns>
        public static TContext FindEnumValue<TContext>(string str)
        {
            TContext result = default;

            string[] enums = Enum.GetNames(typeof(TContext));

            for (int i = 0; i < enums.Length; i++)
            {
                if (string.Compare(str, enums[i], true) == 0)
                {
                    result = (TContext)Enum.ToObject(typeof(TContext), i);
                    break;
                }
            }

            return result;
        }

        /// <summary>
        /// To find Enum Index by string.
        /// </summary>
        /// <typeparam name="TContext">Enum Type</typeparam>
        /// <param name="str">Enum item string</param>
        /// <returns></returns>
        public static int FindEnumIndex<TContext>(string str)
        {
            int result = -1;

            string[] enums = Enum.GetNames(typeof(TContext));

            for (int i = 0; i < enums.Length; i++)
            {
                if (string.Compare(str, enums[i], true) == 0)
                {
                    result = i;
                    break;
                }
            }

            return result;
        }

        /// <summary>
        /// Get Enum value string list.
        /// </summary>
        /// <typeparam name="TContext">Enum Type</typeparam>
        /// <returns></returns>
        public static List<string> GetEnumStringList<TContext>()
        {
            List<string> result = new List<string>();

            string[] enums = Enum.GetNames(typeof(TContext));

            for (int i = 0; i < enums.Length; i++)
            {
                if (result.Contains(enums[i]) == false)
                {
                    result.Add(enums[i]);
                }
            }

            return result;
        }
        /// <summary>
        /// Get Enum value list.
        /// </summary>
        /// <typeparam name="TContext">Enum Type</typeparam>
        /// <returns></returns>
        public static List<TContext> ConvertEnumToList<TContext>()
        {
            List<TContext> result = new List<TContext>();

            string[] enums = Enum.GetNames(typeof(TContext));

            for (int i = 0; i < enums.Length; i++)
            {
                TContext value = (TContext)Enum.ToObject(typeof(TContext), i);
                if (value != null)
                {
                    if (result.Contains(value) == false)
                    {
                        result.Add(value);
                    }
                }
            }

            return result;
        }
        /// <summary>
        /// Get Enum value.
        /// </summary>
        /// <typeparam name="TContext">Enum Type</typeparam>
        /// <returns></returns>
        public static TContext StringToEnum<TContext>(string value)
        {
            return (TContext)Enum.Parse(typeof(TContext), value, true);
        }
        /// <summary>
        /// Get Enum String Value.
        /// </summary>
        /// <typeparam name="TContext">Enum Type</typeparam>
        /// <returns></returns>
        public static string? EnumToString<TContext>(TContext value)
        {
            return value?.ToString();
        }
    }
}
