using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using OSIsoft.AF;
using OSIsoft.AF.Asset;
using OSIsoft.AF.Data;
using OSIsoft.AF.PI;
using OSIsoft.AF.Time;

namespace AFWrapper
{
	[ComVisible(true)]
	public enum WAFUpdateOption
	{
		Replace = 0,
		Insert = 1,
		NoReplace = 2,
		ReplaceOnly = 3,
		InsertNoCompression = 5,
		Remove = 6
	}

	[ComVisible(true)]
	[Guid("A8B1CBE6-7756-49A1-9245-DE1FE7B31AA2")]
	[ClassInterface(ClassInterfaceType.AutoDual)]
	public class WAFErrors 
	{
		private string[,] m_genericErrs;

		public WAFErrors()
		{
			m_genericErrs = new string[0,0];
		}
		public WAFErrors(AFErrors<AFValue> valErrs)
		{
			if (valErrs == null)
			{
				m_genericErrs = new string[0, 0];
			} 
			else
			{
				m_genericErrs = new string[valErrs.Errors.Count, 1];
				int key = 0;

				foreach (AFValue item in valErrs.Errors.Keys)
				{
					m_genericErrs[key, 0] = item.Timestamp.ToString();
					m_genericErrs[key, 1] = valErrs.Errors[item].Message;
					key++;
				}
			}
		}

		[ComVisible(true)]
		public bool HasErrors()
		{
			return(m_genericErrs.Length > 0);
		}

		private string ExceptionToString(Exception ex)
		{
			return (ex.Message);
		}

		[ComVisible(true)]
		public string[,] Errors()
		{
			return(m_genericErrs);
		}

		public override string ToString()
		{
			StringBuilder output = new StringBuilder("None");

			if (this.HasErrors())
			{
				output.Clear();
				for (int i = 0; i < (m_genericErrs.Length / 2); i++)
				{
					output.Append(String.Format("Value at {0} has error: {1}\n\n", m_genericErrs[i,0], m_genericErrs[i,1]));
				}

			}

			return output.ToString();
		}
		
	}


	[ComVisible(true)]
	[Guid("60FFC0D0-669C-4D18-857A-1EA969EFE2A7")]
	[ClassInterface(ClassInterfaceType.AutoDual)]
  public class WAFElement
  {
		private PISystem m_piSys;
		private AFDatabase m_afDb;
		private AFElement m_Element;

		public WAFElement()
		{
			PISystems syss = new PISystems();
			m_piSys = syss.DefaultPISystem;
			m_afDb = m_piSys.Databases.DefaultDatabase;
		}

		[ComVisible(true)]
		public string PISysName
		{
			get
			{
				return (m_piSys.Name);
			}
			set
			{
				try
				{
					PISystems syss = new PISystems();
					m_piSys = syss[value];

				}
				catch (Exception ex)
				{
					System.Console.WriteLine(ex.Message);
				}

			}
		}

		[ComVisible(true)]
		public string setElement(string elemName)
		{
			try
			{
				m_Element = m_afDb.Elements[elemName];
				return (m_Element.Name);
			}
			catch (Exception)
			{
				return ("");
			}
		}
  }


}
