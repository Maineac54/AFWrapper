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
	public enum WAFBoundryType 
	{
		Inside = 0,
		Outside = 1,
		Interpolated
	}

	[ComVisible(true)]
	[Guid("424197AC-DD82-4E11-92CC-8FA8F9700C63")]
	[ClassInterface(ClassInterfaceType.AutoDual)]
	public class WAFTime
	{
		private AFTime m_Time;

		public WAFTime()
		{
			m_Time = new AFTime(DateTime.Now);

		}

		[ComVisible(true)]
		public void setTime(string AnyTime)
		{

			if (!AFTime.TryParse(AnyTime, out m_Time))
			{
				System.Console.WriteLine("Bad input time value.");
				m_Time = new AFTime();  // Empty Time
			}
		}

		// local time
		[ComVisible(true)]
		public DateTime LocalTime {
			get 
			{
				return (m_Time.LocalTime);
			}
			set
			{
				m_Time = new AFTime(value);
			}
		}

		// To String
		[ComVisible(true)]
		public override string ToString()
		{
			return (m_Time.ToString());
		}
		[ComVisible(true)]
		public string ToString(IFormatProvider provider)
		{
			return (m_Time.ToString(provider));
		}
		[ComVisible(true)]
		public string ToString(string format, IFormatProvider provider = null)
		{
			return (m_Time.ToString(format, provider));
		}

		// Equals
		[ComVisible(true)]
		public override bool Equals(object timeObject)
		{
			AFTime test = (AFTime)timeObject;

			return m_Time.Equals(test);
		}
		[ComVisible(true)]
		public override int GetHashCode()
		{
			return m_Time.GetHashCode();
		}

	}

	[ComVisible(true)]
	[Guid("185DDC38-0B52-41B9-92A6-36528E5CBFD8")]
	[ClassInterface(ClassInterfaceType.AutoDual)]
	public class WAFValue
	{
		private AFValue m_AFVal;

		public WAFValue()
		{
			m_AFVal = null;

		}

		[ComVisible(true)]
		public void setAFValue(object val)
		{
			try
			{
				m_AFVal = (AFValue)val;
			}
			catch (Exception)
			{

			}
		}

		[ComVisible(true)]
		public void setAFValue(string Timestamp, string Value)
		{
			AFTime aftTime;

			if (AFTime.TryParse(Timestamp, out aftTime))
			{
				m_AFVal = new AFValue(Value, aftTime);
			}
		}

		[ComVisible(true)]
		public void setAFValue(string Timestamp, string Value, string TagName, string PIServerName)
		{
			AFTime aftTime;

			if (AFTime.TryParse(Timestamp, out aftTime))
			{
				PIServers svrs = new PIServers();
				PIServer piSrv = svrs[PIServerName];
				PIPoint pipt = PIPoint.FindPIPoint(piSrv, TagName);

				m_AFVal = new AFValue(Value, aftTime);
				m_AFVal.PIPoint = pipt;
				
			}
		}

		[ComVisible(true)]
		public object Value
		{
			get
			{
				try
				{
					return (m_AFVal.Value);
				}
				catch (Exception)
				{
					return null;
				}
			}
		}

		[ComVisible(true)]
		public string Timestamp
		{
			get
			{
				return (m_AFVal.Timestamp.LocalTime.ToString());
			}
		}

		[ComVisible(true)]
		public string Status
		{
			get
			{
				return (m_AFVal.Status.ToString());
			}
		}

		public static WAFValue AFValueToWAFValue(AFValue val)
		{
			WAFValue cVal = new WAFValue();

			try
			{
				cVal.setAFValue(val);
			}
			catch (Exception ex)
			{
				System.Console.WriteLine(ex.Message);
				cVal = null;
			}

			return (cVal);
		}

	}

	[ComVisible(true)]
	[Guid("81501521-D440-4111-B30D-A3ED4C950132")]
	[ClassInterface(ClassInterfaceType.AutoDual)]
	public class WPIPoint
	{
		private PIServer m_piSrv;
		private PIPoint m_Pipt;

		public WPIPoint()
		{
			PIServers srvrs = new PIServers();
			m_piSrv = srvrs.DefaultPIServer;

		}

		[ComVisible(true)]
		public string PIServerName
		{
			get
			{
				return (m_piSrv.Name);
			}
		}

		// Set the PI Point name and get the point
		[ComVisible(true)]
		public string Name
		{
			set
			{
				try
				{
					m_Pipt = PIPoint.FindPIPoint(m_piSrv, value);
				}
				catch (Exception ex)
				{
					System.Console.WriteLine(ex.Message);
					m_Pipt = null;
				}
			}
			get
			{
				try
				{
					return (m_Pipt.Name);
				}
				catch (Exception)
				{
					return ("");
				}
			}
		}

		[ComVisible(true)]
		public string setHost(string HostName)
		{
			AFConnectionPreference pref = m_piSrv.ConnectionInfo.Preference;

			PICollective pCol = m_piSrv.Collective;
			PICollectiveMember mbr = pCol.Members[HostName];
			mbr.Connect();

			return (pCol.CurrentMember.Name);
		}

		// Data Methods
		// Snapshot
		[ComVisible(true)]
		public WAFValue Snapshot()
		{
			try
			{
				AFValue val = m_Pipt.Snapshot();
				WAFValue wVal = new WAFValue();
				wVal.setAFValue(val);

				return (wVal);
			}
			catch (Exception)
			{
				return (null);
			}
		}

		// Sampled Values
		[ComVisible(true)]
		public void InterpolatedValues(WAFTime start, WAFTime end, string interval, string filterExp, bool includeFiltVals, ref WAFValue[] values)
		{
			List<WAFValue> retVals = null;

			try
			{
				AFTimeRange range = new AFTimeRange(start.ToString(), end.ToString());
				
				AFTimeSpan span = new AFTimeSpan();
				if (!AFTimeSpan.TryParse(interval, out span))
				{
					span = AFTimeSpan.Parse("1h");
				}

				AFValues vals = m_Pipt.InterpolatedValues(range, span, filterExp, includeFiltVals);

				retVals = vals.ConvertAll(new Converter<AFValue, WAFValue>(WAFValue.AFValueToWAFValue));
			}
			catch (Exception ex)
			{
				
				System.Console.WriteLine(ex.Message);
			}

			values = retVals.ToArray();
			return;
		}

		// Recorded Values
		[ComVisible(true)]
		public void RecordedValues(WAFTime start, WAFTime end, WAFBoundryType boundryType, string filterExp, bool includeFiltVals, int maxCount, ref WAFValue[] output)
		{
			List<WAFValue> retVals = null;

			try
			{
				AFTimeRange range = new AFTimeRange(start.ToString(), end.ToString());
				AFBoundaryType opt = (AFBoundaryType)Enum.ToObject(typeof(AFBoundaryType), (byte)boundryType);

				AFValues vals = m_Pipt.RecordedValues(range, AFBoundaryType.Interpolated, filterExp, includeFiltVals, maxCount);

				retVals = vals.ConvertAll(new Converter<AFValue, WAFValue>(WAFValue.AFValueToWAFValue));
			}
			catch (Exception ex)
			{
				System.Console.WriteLine(ex.Message);
			}

			output = retVals.ToArray();
			return;
		}

		// Update Value
		[ComVisible(true)]
		public WAFErrors UpdateValues(ref WAFValue[] vals, WAFUpdateOption UpdateOption)
		{
			List<WAFValue> wafVals = vals.ToList<WAFValue>();

			List<AFValue> afVals = new List<AFValue>();
			foreach (WAFValue wafVal in wafVals)
			{
				AFValue afVal = m_Pipt.RecordedValue(new AFTime(wafVal.Timestamp), AFRetrievalMode.Exact);
				afVals.Add(afVal);
			}

			AFUpdateOption opt =  (AFUpdateOption)Enum.ToObject(typeof(AFUpdateOption), (byte)UpdateOption);

			AFErrors<AFValue> errs = m_Pipt.UpdateValues(afVals, opt, AFBufferOption.Buffer);

			WAFErrors retVal = new WAFErrors(errs);
			return (retVal);
		}

	}
}
