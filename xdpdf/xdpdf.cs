/******************************************************************************
 *  copyright            : 2011 by Daniel Brakensiek                          *
 *  email                : daniel@braktech.com                                *
 ******************************************************************************/

/******************************************************************************
 *   This file is part of xdpdf.                                              *
 *                                                                            *
 *   xdpdf is free software: you can redistribute it and/or modify it under   *
 *   the terms of the GNU Lesser General Public License as published by the   *
 *   Free Software Foundation, either version 3 of the License, or (at your   *
 *   option) any later version.                                               *
 *                                                                            *
 *   xdpdf is distributed in the hope that it will be useful, but WITHOUT     *
 *   ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or    *
 *   FITNESS FOR A PARTICULAR PURPOSE.  See the GNU Lesser General Public     *
 *   License for more details.                                                *
 *                                                                            *
 *   You should have received a copy of the GNU Lesser General Public         *
 *   License along with xdpdf.  If not, see <http://www.gnu.org/licenses/>.   *
 ******************************************************************************/

#define TRACE

using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.IO;
using System.Diagnostics;
using System.Configuration;
using Microsoft.Exchange.Data.Transport;
using Microsoft.Exchange.Data.Mime;
using Microsoft.Exchange.Data.Transport.Email;
using Microsoft.Exchange.Data.Transport.Smtp;
using Microsoft.Exchange.Data.Transport.Routing;
using Microsoft.Exchange.Data.Common;
using Microsoft.Exchange.Data.TextConverters;

[assembly: CLSCompliant(true)]
namespace xdpdf
{
    public sealed class xdpdfFactory : RoutingAgentFactory
    {
        public override RoutingAgent CreateAgent(SmtpServer server)
        {
            RoutingAgent routingagent = new xdpdfRoutingAgent();
            return routingagent;
        }
    }

    public class xdpdfRoutingAgent : RoutingAgent
    {
        public xdpdfRoutingAgent()
        {
            base.OnSubmittedMessage += new SubmittedMessageEventHandler(xdpdfRoutingAgent_OnSubmittedMessage);
        }

        void xdpdfRoutingAgent_OnSubmittedMessage(SubmittedMessageEventSource source, QueuedMessageEventArgs e)
        {
            Stopwatch stopwatch = new Stopwatch();
            String messageid = e.MailItem.Message.MessageId.Substring(1, e.MailItem.Message.MessageId.Length - 2); // Strip brackets from messageid
            Log log = Log.Instance;
            log.Do(messageid + ": starting process", 1);
            stopwatch.Start();
            //Detect a PDF among attachments
            for (int i = e.MailItem.Message.Attachments.Count - 1; i >= 0; i--)
            {
                Attachment attachment = e.MailItem.Message.Attachments[i];
                if (xdpdf.xdpdfSettings.Default.ScanAllAttachments == true || attachment.ContentType == "application/pdf" || attachment.FileName.Substring(attachment.FileName.Length - 3, 3).ToLower() == "pdf")
                {
                    log.Do(messageid + ": processing attachment: \"" + attachment.FileName + "\"", 2);
                    Stream attachreadstream = attachment.GetContentReadStream();
                    PDFTools tools = new PDFTools(messageid, e.MailItem.Message.Attachments[i].FileName, attachreadstream);
                    if (tools.Detect() == true)
                    {
                        log.Do(messageid + ": attachment \"" + attachment.FileName + "\" is detected as a PDF", 2);
                        if (tools.ScanPDF())
                        {
                            String footerstring = "The PDF attachment " + attachment.FileName + " has been disarmed. ";
                            footerstring += "If it no longer works, please forward the following information to your mail administrator:";
                            PDFTools.AddFooterToBody(messageid, e.MailItem.Message.Body, footerstring);
                            footerstring = messageid + "::" + tools.AttachGuid;
                            PDFTools.AddFooterToBody(messageid, e.MailItem.Message.Body, footerstring);
                            Stream attachwritestream = attachment.GetContentWriteStream();
                            tools.DisarmedStream.WriteTo(attachwritestream);
                            attachwritestream.Flush();
                            attachwritestream.Close();
                        }
                    }
                    else
                    {
                        log.Do(messageid + ": attachment \"" + attachment.FileName + "\" is not detected as a PDF", 2);
                    }
                    attachreadstream.Close();
                }
            }
            stopwatch.Stop();
            log.Do(messageid + ": finished - processing took " + stopwatch.Elapsed.Milliseconds + "ms", 1);
            Trace.Flush();
        }
    }

    public sealed class Log : IDisposable
    {
        private static readonly Log instance = new Log();
        private Boolean logging;
        private Int32 desiredloglevel;
        private String logpath;
        private String currentstamp;
        private Object lockable = new object();
        private TextWriterTraceListener textlistener;
        private StreamWriter logfile;

        private Log()
        {
            if (xdpdf.xdpdfSettings.Default.Logging)
            {
                logging = true;
                currentstamp = DateTime.Now.ToString("yyyyMMdd");
                logpath = xdpdf.xdpdfSettings.Default.LogPath;
                AttachNewListener(currentstamp);
            }
            else
            {
                logging = false;
            }
            if (xdpdf.xdpdfSettings.Default.LogLevel > -1 && xdpdf.xdpdfSettings.Default.LogLevel < 3)
            {
                desiredloglevel = xdpdf.xdpdfSettings.Default.LogLevel;
            }
            else
            {
                desiredloglevel = 2;
            }
        }

        public void Dispose()
        {
            logfile.Dispose();
            textlistener.Dispose();
            GC.SuppressFinalize(this);
        }

        private void RemoveOldListener()
        {
            Trace.Listeners.Remove(textlistener);
            textlistener.Close();
            logfile.Close();
        }

        private void AttachNewListener(String stamp)
        {
            if (logpath.Substring(logpath.Length - 1, 1) != "\\")
            {
                logpath += "\\";
            }
            logfile = File.AppendText(logpath + stamp + "-disarm.log");
            textlistener = new TextWriterTraceListener(logfile);
            Trace.Listeners.Add(textlistener);
        }

        public void Do(String msg, Int32 level)
        {
            if (!logging) // Not logging anything
            {
                return;
            }
            if (level > desiredloglevel) // Only logging entries at our specified level
            {
                return;
            }
            string newstamp = DateTime.Now.ToString("yyyyMMdd");
            if (currentstamp != newstamp)
            {
                lock (lockable)
                {
                    if (currentstamp != newstamp)
                    {
                        RemoveOldListener();
                        AttachNewListener(newstamp);
                        currentstamp = newstamp;
                    }
                }
            }
            msg = DateTime.Now.ToString("yyyyMMddHHmmssffff") + ": " + msg;
            Trace.WriteLine(msg);
        }

        public static Log Instance
        {
            get
            {
                return instance;
            }
        }
    }

    public class PDFTools : IDisposable
    {
        String messageid;
        Stream attachstream;
        MemoryStream disarmedstream;
        String attachguid;
        String attachname;
        Dictionary<int, String> disarmbytes = new Dictionary<int, String>();
        Boolean isdisarmed = false;
        int longestword = 0;
        byte[] original;
        byte[] disarmed;
        byte[] header_magic = {
                              0x25,
                              0x50,
                              0x44,
                              0x46
                          }; // %PDF
        String[] Keywords;
        String outpath;
        List<String> detectedwords = new List<String>();

        public PDFTools(String id, String name, Stream instream)
        {
            outpath = xdpdf.xdpdfSettings.Default.QuarantinePath;
            if (outpath.Substring(outpath.Length - 1, 1) != "\\")
            {
                outpath += "\\";
            }
            Keywords = new String[xdpdf.xdpdfSettings.Default.Keywords.Count];
            xdpdf.xdpdfSettings.Default.Keywords.CopyTo(Keywords, 0);
            messageid = id;
            attachname = name;
            attachstream = instream;
            foreach (String keyword in Keywords)
            {
                if (keyword.Length * 3 > longestword)
                {
                    longestword = keyword.Length * 3; // Need to allow for any word to be completely hex encoded
                }
            }
        }

        public String AttachGuid
        {
            get { return attachguid; }
            set { attachguid = value; }
        }

        public MemoryStream DisarmedStream
        {
            get { return disarmedstream; }
            set { disarmedstream = value; }
        }

        public Boolean Detect()
        {
            byte[] b = new byte[1024];
            attachstream.Read(b, 0, 1024);
            return CheckForSequence(b, header_magic);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (disposing)
            {
                disarmedstream.Dispose();
            }
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        public Boolean ScanPDF()
        {
            int streamlength = Convert.ToInt32(attachstream.Length);
            int streamposition = 0;
            Boolean inword = false;
            String character = "";
            String word = "";
            int wordlength = 0;
            List<String> wordexact = new List<String>();
            ASCIIEncoding encoding = new ASCIIEncoding();
            original = new byte[streamlength];
            disarmed = new byte[streamlength];
            String digit1 = "";
            String digit2 = "";
            String hexchar;
            int hexint;
            Boolean detected;
            int disarmposition;
            Log log = Log.Instance;

            original = ReadToEnd(attachstream, streamlength);
            Collection<int> slashes = FindBytes(original, 0x2f); // Searching for '/'
            foreach (int slashlocation in slashes)
            {
                streamposition = slashlocation + 1; // Start of the word
                int streamoffset = 0;
                inword = true;
                word = "";
                wordexact = new List<string>();
                wordlength = 0;

                while (streamoffset < longestword && inword && streamposition + streamoffset < streamlength)
                {
                    character = encoding.GetString(original, streamposition + streamoffset, 1);
                    if (IsAlphaNumeric(character))
                    {
                        word += character;
                        wordexact.Add(character);
                        wordlength++;
                    }
                    else if (character.Equals("#") && streamposition + streamoffset < (streamlength - 2))
                    {
                        digit1 = encoding.GetString(original, streamposition + streamoffset + 1, 1);
                        digit2 = encoding.GetString(original, streamposition + streamoffset + 2, 1);
                        if (IsHexadecimal(digit1) && IsHexadecimal(digit2))
                        {
                            hexchar = digit1 + digit2;
                            hexint = Convert.ToInt32(hexchar, 16);
                            character = Convert.ToChar(hexint).ToString();
                            word += character;
                            wordexact.Add(hexchar);
                            wordlength += 3;
                            streamoffset += 2;
                        }
                        else
                        {
                            inword = false;
                        }
                    }
                    else
                    {
                        inword = false;
                    }
                    streamoffset++;
                }
                // The word has finished; now disarm
                detected = UpdateWords(word, wordexact);
                if (detected)
                {
                    isdisarmed = true;
                    disarmposition = slashlocation + 1;
                    disarmbytes[disarmposition] = "x";
                    disarmbytes[disarmposition + 1] = "x";
                }
            }

            if (isdisarmed)
            {
                log.Do(messageid + ": attachment \"" + attachname + "\" is being disarmed", 0);
                LogAndDisarm();
            }
            else
            {
                log.Do(messageid + ": attachment \"" + attachname + "\" did not contain any specified keywords", 1);
            }
            return isdisarmed;
        }

        private Boolean UpdateWords(string word, List<string> word_exact)
        {
            Boolean detected = false;
            if (Keywords.Contains("/" + word))
            {
                detected = true;
                if (!detectedwords.Contains(word))
                {
                    detectedwords.Add(word);
                }
            }
            return detected;
        }

        void LogAndDisarm()
        {
            Log log = Log.Instance;
            original.CopyTo(disarmed, 0);

            string suspect_words = "";
            foreach (var item in detectedwords)
            {
                suspect_words += item + ", ";
            }
            suspect_words = suspect_words.Substring(0, suspect_words.Length - 2);
            log.Do(messageid + ": attachment \"" + attachname + "\" contained the following suspicious keywords: " + suspect_words, 1);
            
            foreach (var location in disarmbytes)
            {
                disarmed[location.Key] = Encoding.ASCII.GetBytes(location.Value)[0];
            }

            String fullpath = outpath + messageid;
            Directory.CreateDirectory(fullpath);
            attachguid = System.Guid.NewGuid().ToString();
            fullpath = fullpath + "\\" + attachguid;
            using (FileStream rawstream = new FileStream(fullpath, FileMode.CreateNew))
            {
                BinaryWriter rawstreamwriter = new BinaryWriter(rawstream);
                rawstreamwriter.Write(original);
            }
            
            disarmedstream = new MemoryStream();
            BinaryWriter disarmedstreamwriter = new BinaryWriter(disarmedstream);
            disarmedstreamwriter.Write(disarmed);
            
            return;
        }

        static private Boolean IsAlphaNumeric(string character)
        {
            if (CharInRange(character, "A", "Z") || CharInRange(character, "0", "9"))
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        static private Boolean IsHexadecimal(string character)
        {
            if (CharInRange(character, "A", "F") || CharInRange(character, "0", "9"))
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        static private Boolean CharInRange(String testchar, String firstchar, String secondchar)
        {
            if (String.Compare(testchar, firstchar, true) >= 0 && String.Compare(testchar, secondchar, true) <= 0)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        static private byte[] ReadToEnd(Stream stream, int length)
        {
            long originalposition = stream.Position;
            stream.Position = 0;

            if (length < 1)
            {
                length = 32768;
            }

            byte[] buffer = new byte[length];
            int read = 0;

            int chunk;
            while ((chunk = stream.Read(buffer, read, buffer.Length - read)) > 0)
            {
                read += chunk;

                //If we're at the end of the buffer, check for more data
                if (read == buffer.Length)
                {
                    int nextbyte = stream.ReadByte();
                    // if the next byte is -1 we're at the end of the stream
                    if (nextbyte == -1)
                    {
                        stream.Position = originalposition;
                        return buffer;
                    }
                    // Still more data; keep reading
                    byte[] newbuffer = new byte[buffer.Length * 2];
                    Array.Copy(buffer, newbuffer, buffer.Length);
                    buffer = newbuffer;
                    read++;
                }
            }
            //The buffer is probably too big - shrink it before returning
            byte[] ret = new byte[read];
            Array.Copy(buffer, ret, read);
            stream.Position = originalposition;
            return ret;
        }

        public static Boolean CheckForSequence(byte[] buffer, byte[] pattern)
        {
            Boolean found = false;
            int i = Array.IndexOf<byte>(buffer, pattern[0], 0);
            while (i >= 0 && i <= buffer.Length - pattern.Length && !found)
            {
                byte[] segment = new byte[pattern.Length];
                Buffer.BlockCopy(buffer, i, segment, 0, pattern.Length);
                if (segment.SequenceEqual<byte>(pattern))
                {
                    found = true;
                }
                i = Array.IndexOf<byte>(buffer, pattern[0], i + pattern.Length);
            }
            return found;
        }

        public static Collection<int> FindBytes(byte[] buffer, byte searchfor)
        {
            Collection<int> positions = new Collection<int>();
            int i = Array.IndexOf<byte>(buffer, searchfor, 0);
            while (i >= 0 && i <= buffer.Length - 1)
            {
                positions.Add(i);
                i = Array.IndexOf<byte>(buffer, searchfor, i + 1);
            }
            return positions;
        }

        static public void AddFooterToBody(String messageid, Microsoft.Exchange.Data.Transport.Email.Body body, String text)
        {
            Stream originalbodycontent = null;
            Stream newbodycontent = null;
            Encoding encoding;
            String charsetname;
            Log log = Log.Instance;

            try
            {
                BodyFormat bodyformat = body.BodyFormat;
                if (!body.TryGetContentReadStream(out originalbodycontent))
                {
                    //body can't be decoded
                    log.Do(messageid + ": email body format could not be decoded - warning footer not appended", 0);
                }
                if (BodyFormat.Text == bodyformat)
                {
                    charsetname = body.CharsetName;
                    if (null == charsetname || !Microsoft.Exchange.Data.Globalization.Charset.TryGetEncoding(charsetname, out encoding))
                    {
                        // either no charset, or charset is not supported by the system
                        log.Do(messageid + ": email body character set is either not defined or not supported by the system - warning footer not appended", 0);
                    }
                    else
                    {
                        TextToText texttotextconversion = new TextToText();
                        texttotextconversion.InputEncoding = encoding;
                        texttotextconversion.HeaderFooterFormat = HeaderFooterFormat.Text;
                        texttotextconversion.Footer = text;
                        newbodycontent = body.GetContentWriteStream();
                        try
                        {
                            texttotextconversion.Convert(originalbodycontent, newbodycontent);
                        }
                        catch (Microsoft.Exchange.Data.TextConverters.TextConvertersException)
                        {
                            log.Do(messageid + ": error while performing body text conversion - warning footer not appended", 0);
                        }
                    }

                }
                else if (BodyFormat.Html == bodyformat)
                {
                    charsetname = body.CharsetName;
                    if (null == charsetname ||
                        !Microsoft.Exchange.Data.Globalization.Charset.TryGetEncoding(charsetname, out encoding))
                    {
                        log.Do(messageid + ": email body character set is either not defined or unsupported - warning footer not appended", 0);
                    }
                    else
                    {
                        HtmlToHtml htmltohtmlconversion = new HtmlToHtml();
                        htmltohtmlconversion.InputEncoding = encoding;
                        htmltohtmlconversion.HeaderFooterFormat = HeaderFooterFormat.Html;
                        htmltohtmlconversion.Footer = "<p><font size=\"-1\">" + text + "</font></p>";
                        newbodycontent = body.GetContentWriteStream();

                        try
                        {
                            htmltohtmlconversion.Convert(originalbodycontent, newbodycontent);
                        }
                        catch (Microsoft.Exchange.Data.TextConverters.TextConvertersException)
                        {
                            // the conversion has failed..
                            log.Do(messageid + ": error while performing body html conversion - warning footer not appended", 0);
                        }
                    }

                }
                else if (BodyFormat.Rtf == bodyformat)
                {
                    RtfToRtf rtftortfconversion = new RtfToRtf();
                    rtftortfconversion.HeaderFooterFormat = HeaderFooterFormat.Html;
                    rtftortfconversion.Footer = "<font face=\"Arial\" size=\"+1\">" + text + "</font>";
                    Stream uncompressedbodycontent = body.GetContentWriteStream();

                    try
                    {
                        rtftortfconversion.Convert(originalbodycontent, uncompressedbodycontent);
                    }
                    catch (Microsoft.Exchange.Data.TextConverters.TextConvertersException)
                    {
                        //Conversion failed
                        log.Do(messageid + ": error while decompressing body rtf - warning footer not appended", 0);
                    }

                    RtfToRtfCompressed rtfcompressionconversion = new RtfToRtfCompressed();
                    rtfcompressionconversion.CompressionMode = RtfCompressionMode.Compressed;
                    newbodycontent = body.GetContentWriteStream();

                    try
                    {
                        rtfcompressionconversion.Convert(uncompressedbodycontent, newbodycontent);
                    }
                    catch (Microsoft.Exchange.Data.TextConverters.TextConvertersException)
                    {
                        // the conversion has failed..
                        log.Do(messageid + ": error compressing body rtf - warning footer not appended", 0);
                    }
                }

                else
                {
                    // Handle cases where the body format is not one of the above.
                    log.Do(messageid + ": unsupported body email format - warning footer not appended", 0);
                }
            }

            finally
            {
                if (originalbodycontent != null)
                {
                    originalbodycontent.Close();
                }

                if (newbodycontent != null)
                {
                    newbodycontent.Close();
                }
            }
        }
    }
}
