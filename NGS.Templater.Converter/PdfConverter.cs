using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using Microsoft.Win32;
using unoidl.com.sun.star.frame;
using unoidl.com.sun.star.lang;
using unoidl.com.sun.star.util;

namespace NGS.Templater
{
	/// <summary>
	/// Utility for converting various document formats to pdf.
	/// Uses OpenOffice.org/LibreOffice to make conversion
	/// </summary>
	public static class PdfConverter
	{
		public static void Main(string[] args)
		{
			if (args.Length == 2)
				PdfConverter.Convert(args[0], args[1]);
			else
				Console.WriteLine("Usage: PdfConverter.exe fromFile toFile");
		}

		private static XComponentLoader aLoader;
		private static XComponentLoader Loader
		{
			get
			{
				if (aLoader == null)
					Init();
				return aLoader;
			}
		}

		private static void Init()
		{
			aLoader = InitLoader();
			if (aLoader == null && InitOpenOffice3Environment())
				aLoader = InitLoader();
			if (aLoader == null && InitLibreOfficeEnvironment())
				aLoader = InitLoader();
			if (aLoader == null)
				throw new MissingMethodException("Can't find OpenOffice.org or LibreOffice. Office must be installed for pdf conversion.");
		}

		private static XComponentLoader InitLoader()
		{
			try
			{
				XComponentLoader loader = null;
				var thread = new Thread(() =>
				{
					try
					{
						var xLocalContext = uno.util.Bootstrap.bootstrap();
						var xRemoteFactory = (unoidl.com.sun.star.lang.XMultiServiceFactory)xLocalContext.getServiceManager();
						loader = (XComponentLoader)xRemoteFactory.createInstance("com.sun.star.frame.Desktop");
					}
					catch { }
				});
				thread.Start();
				thread.Join(TimeSpan.FromMinutes(1));
				return loader;
			}
			catch
			{
				return null;
			}
		}

		private static bool InitOpenOffice3Environment()
		{
			try
			{
				//from http://blog.nkadesign.com/2008/net-working-with-openoffice-3/ 
				string baseKey = null;
				var baseKey64 = "SOFTWARE\\Wow6432Node\\OpenOffice.org\\";
				var baseKey32 = "SOFTWARE\\OpenOffice.org\\";
				// OpenOffice being a 32 bit app, its registry location is different in a 64 bit OS 
				if ((Marshal.SizeOf(typeof(IntPtr)) == 8))
					baseKey = baseKey64;
				else
					baseKey = baseKey32;

				// Get the URE directory 
				string key = (baseKey + "Layers\\URE\\1");
				var reg = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(key);
				if ((reg == null))
					reg = Registry.LocalMachine.OpenSubKey(key);

				if (reg == null)
				{
					if (baseKey == baseKey32)
					{
						key = (baseKey64 + "Layers\\URE\\1");
						reg = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(key);
						if (reg == null)
							reg = Registry.LocalMachine.OpenSubKey(key);
						if (reg == null)
							return false;
					}
					else
						return false;
				}

				var urePath = (string)reg.GetValue("UREINSTALLLOCATION");
				reg.Close();
				urePath = System.IO.Path.Combine(urePath, "bin");
				// Get the UNO Path 
				key = (baseKey + "UNO\\InstallPath");
				reg = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(key);
				if ((reg == null))
					reg = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(key);

				var unoPath = (string)reg.GetValue(null);
				reg.Close();

				var path = string.Format("{0};{1}", System.Environment.GetEnvironmentVariable("PATH"), urePath);
				Environment.SetEnvironmentVariable("PATH", path);
				Environment.SetEnvironmentVariable("UNO_PATH", unoPath);

				return true;
			}
			catch
			{
				return false;
			}
		}

		private static bool InitLibreOfficeEnvironment()
		{
			try
			{
				string baseKey = null;
				var baseKey64 = "SOFTWARE\\Wow6432Node\\LibreOffice\\";
				var baseKey32 = "SOFTWARE\\LibreOffice\\";
				// OpenOffice being a 32 bit app, its registry location is different in a 64 bit OS 
				if ((Marshal.SizeOf(typeof(IntPtr)) == 8))
					baseKey = baseKey64;
				else
					baseKey = baseKey32;

				// Get the URE directory 
				var key = (baseKey + "Layers_\\URE\\1");
				var reg = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(key);
				if ((reg == null))
					reg = Registry.LocalMachine.OpenSubKey(key);

				if (reg == null)
				{
					if (baseKey == baseKey32)
					{
						key = (baseKey64 + "Layers_\\URE\\1");
						reg = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(key);
						if (reg == null)
							reg = Registry.LocalMachine.OpenSubKey(key);
						if (reg == null)
							return false;
					}
					else
						return false;
				}

				var urePath = (string)reg.GetValue("UREINSTALLLOCATION");
				reg.Close();
				urePath = System.IO.Path.Combine(urePath, "bin");
				// Get the UNO Path 
				key = (baseKey + "UNO\\InstallPath");
				reg = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(key);
				if ((reg == null))
					reg = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(key);

				var unoPath = (string)reg.GetValue(null);
				reg.Close();

				var path = string.Format("{0};{1}", Environment.GetEnvironmentVariable("PATH"), urePath);
				Environment.SetEnvironmentVariable("PATH", path);
				Environment.SetEnvironmentVariable("UNO_PATH", unoPath);

				return true;
			}
			catch
			{
				return false;
			}
		}

		/// <summary>
		/// Converts input document (ex. MyDocument.xlsx) to output pdf
		/// </summary>
		/// <param name="content">input document content</param>
		/// <param name="ext">input document extension</param>
		/// <returns>output pdf content</returns>
		public static byte[] Convert(byte[] content, string ext)
		{
			var fajl = Path.GetTempPath() + "Input" + Guid.NewGuid().ToString() + "." + ext;
			var pdf = Path.GetTempPath() + "Output" + Guid.NewGuid().ToString() + ".pdf";
			File.WriteAllBytes(fajl, content);
			Convert(fajl, pdf);
			var result = File.ReadAllBytes(pdf);
			File.Delete(fajl);
			File.Delete(pdf);
			return result;
		}

		private static object sync = new object();

		/// <summary>
		/// Converts input document (ex. MyDocument.docx) to output pdf
		/// Requires full path to document (ex. C:\Documents\MyDocument.docx)
		/// </summary>
		/// <param name="from">input document</param>
		/// <param name="toPdf">output pdf file</param>
		public static void Convert(string from, string toPdf)
		{
			if (!File.Exists(from))
				throw new FileNotFoundException("Can't find input file", from);
			var pv = new unoidl.com.sun.star.beans.PropertyValue[1];
			pv[0] = new unoidl.com.sun.star.beans.PropertyValue();
			pv[0].Name = "Hidden";
			pv[0].Value = new uno.Any(true);
			lock (sync)
			{
				XComponent xComponent;
				try
				{
					xComponent = Loader.loadComponentFromURL("file:///" + from.Replace('\\', '/'), "_blank", 0, pv);
				}
				catch (DisposedException)
				{
					Init();
					xComponent = Loader.loadComponentFromURL("file:///" + from.Replace('\\', '/'), "_blank", 0, pv);
				}
				var xStorable = (XStorable)xComponent;
				pv[0].Name = "FilterName";
				switch (Path.GetExtension(from).ToLowerInvariant())
				{
					case ".xls":
					case ".xlsx":
					case ".ods":
						pv[0].Value = new uno.Any("calc_pdf_Export");
						break;
					default:
						pv[0].Value = new uno.Any("writer_pdf_Export");
						break;
				}
				xStorable.storeToURL("file:///" + toPdf.Replace('\\', '/'), pv);
				var xClosable = (XCloseable)xComponent;
				xClosable.close(true);
			}
		}
	}
}
