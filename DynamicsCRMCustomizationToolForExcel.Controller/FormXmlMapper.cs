
using DynamicsCRMCustomizationToolForExcel.Model.FetchXml;
using DynamicsCRMCustomizationToolForExcel.Model.FormXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Serialization;

namespace DynamicsCRMCustomizationToolForExcel.Controller
{
    public class FormXmlMapper
    {
        public static FormType MapFormXmlToObj(string formxml)
        {
            FormType formType = null;
            try
            {
                StringReader stringReader = null;
                using (stringReader = new StringReader(formxml))
                {
                    XmlSerializer serializer1 = new XmlSerializer(typeof(FormType));
                    formType = (FormType)serializer1.Deserialize(stringReader);
                }
            }
            catch (Exception)
            {
                //------To Handle
            }
            return formType;
        }

        public static savedqueryLayoutxmlGrid MapViewXmlToObj(string viewXml)
        {
            savedqueryLayoutxmlGrid viewType = null;
            try
            {
                StringReader stringReader = null;
                using (stringReader = new StringReader(viewXml))
                {
                    XmlSerializer serializer1 = new XmlSerializer(typeof(savedqueryLayoutxmlGrid));
                    viewType = (savedqueryLayoutxmlGrid)serializer1.Deserialize(stringReader);
                }
            }
            catch (Exception e)
            {
                //------To Handle
            }
            return viewType;
        }

        public static FetchType MapFetchXmlToObj(string fetchXml)
        {
            FetchType viewType = null;
            try
            {
                StringReader stringReader = null;
                using (stringReader = new StringReader(fetchXml))
                {
                    XmlSerializer serializer1 = new XmlSerializer(typeof(FetchType));
                    viewType = (FetchType)serializer1.Deserialize(stringReader);
                }
            }
            catch (Exception e)
            {
                //------To Handle
            }
            return viewType;
        }

        public static string MapObjToViewXml(savedqueryLayoutxmlGrid viewXml)
        {
            string viewstring = null;
            try
            {
                XmlWriterSettings settings = new XmlWriterSettings();
                settings.OmitXmlDeclaration = true;
                using (MemoryStream ms = new MemoryStream())
                {
                    XmlSerializerNamespaces names = new XmlSerializerNamespaces();
                    names.Add("", "");
                    XmlWriter stringWriter = null;
                    using (stringWriter = XmlWriter.Create(ms, settings))
                    {
                        XmlSerializer serializer1 = new XmlSerializer(typeof(savedqueryLayoutxmlGrid));
                        serializer1.Serialize(stringWriter, viewXml, names);
                        ms.Flush();
                        ms.Seek(0, SeekOrigin.Begin);
                        using (StreamReader sr = new StreamReader(ms))
                        {
                            viewstring = sr.ReadToEnd();
                        }
                    }
                }
            }
            catch (Exception e)
            {
                //------To Handle
            }
            return viewstring;
        }

        public static string MapObjToFetchXml(FetchType fetchXml)
        {
            string viewstring = null;
            try
            {
                XmlWriterSettings settings = new XmlWriterSettings();
                settings.OmitXmlDeclaration = true;
                using (MemoryStream ms = new MemoryStream())
                {
                    XmlSerializerNamespaces names = new XmlSerializerNamespaces();
                    names.Add("", "");
                    XmlWriter stringWriter = null;
                    using ( stringWriter = XmlWriter.Create(ms, settings))
                    {
                        XmlSerializer serializer1 = new XmlSerializer(typeof(FetchType));
                        serializer1.Serialize(stringWriter, fetchXml,names);
                        ms.Flush();
                        ms.Seek(0, SeekOrigin.Begin);
                        using(StreamReader sr = new StreamReader(ms)){
                            viewstring = sr.ReadToEnd();
                        }
                    }
                }
            }
            catch (Exception e)
            {
                //------To Handle
            }
            return viewstring;
        }

        public static string GetFormXmlLocalizedLabel(int langauge, FormXmlLabelsTypeLabel[] labels)
        {
            int intComparsison;
            IEnumerable<FormXmlLabelsTypeLabel> label = labels.Where(x => int.TryParse(x.languagecode, out intComparsison) && intComparsison == langauge);
            if (label.Count() > 0)
            {
                return label.FirstOrDefault().description ?? string.Empty;
            }
            return string.Empty;
        }

        public static string GetAttributeEvents(FormXmlEventsTypeEvent[] events, string controlId)
        {
            StringBuilder eventString = new StringBuilder();
            bool first = true;
            if (events != null)
            {
                foreach (var evt in events)
                {
                    if (!evt.application && evt.attribute == controlId)
                    {
                        foreach (var handler in evt.Handlers)
                        {
                            if (!first)
                                eventString.Append("\n");
                            else
                                first = false;

                            eventString.Append(string.Format("Library: {1} -Function: {0} - Enabled :{2}", handler.functionName, handler.libraryName, handler.enabled));
                        }
                    }
                }
            }
            return eventString.ToString();
        }

    }
}
