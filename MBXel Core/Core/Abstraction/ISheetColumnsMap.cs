using System.Collections.Generic;

namespace MBXel_Core.Core.Abstraction
{
    public interface ISheetColumnsMap<T> where T : class
    {
        /// <summary>
        /// Create a map from a worksheet columns to <see cref="T"/> properties 
        /// </summary>
        /// <returns>
        /// <see cref="Dictionary{string , string}"/>
        /// <br/>
        /// <br/>
        /// The <b>key</b> <see cref="string"/> represent a <see cref="T"/>'s property name
        /// The <b>value</b> <see cref="string"/> represent a worksheet column name to be mapped
        /// <br/>
        /// <br/>
        /// <b>Syntax:</b>
        /// <br/>
        /// dictionary.Add(nameof(T.Property), "Column name");
        /// <example>
        /// <br/>
        /// <br/>
        /// <b>Example:</b>
        /// <br/>
        ///<code>
        /// map.Add(nameof(Student.Id), "Student ID");
        /// <br/>
        /// map.Add(nameof(Student.BranchId), "Branch"); 
        /// </code>
        /// </example>
        /// </returns>
        Dictionary<string , string> CreateMap();
    }
}
