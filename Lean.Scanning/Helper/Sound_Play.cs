using System;
using System.Collections;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Windows.Forms;



namespace Lean.Scanning
{
    class Sound_Play
    {
        [Flags]
        public enum PlaySoundFlags : int
        {
            SND_SYNC = 0x0000,    /*  play  synchronously  (default)  */  //同步  
            SND_ASYNC = 0x0001,    /*  play  asynchronously  */  //异步  
            SND_NODEFAULT = 0x0002,    /*  silence  (!default)  if  sound  not  found  */
            SND_MEMORY = 0x0004,    /*  pszSound  points  to  a  memory  file  */
            SND_LOOP = 0x0008,    /*  loop  the  sound  until  next  sndPlaySound  */
            SND_NOSTOP = 0x0010,    /*  don't  stop  any  currently  playing  sound  */
            SND_NOWAIT = 0x00002000,  /*  don't  wait  if  the  driver  is  busy  */
            SND_ALIAS = 0x00010000,  /*  name  is  a  registry  alias  */
            SND_ALIAS_ID = 0x00110000,  /*  alias  is  a  predefined  ID  */
            SND_FILENAME = 0x00020000,  /*  name  is  file  name  */
            SND_RESOURCE = 0x00040004    /*  name  is  resource  name  or  atom  */
        }

        [DllImport("winmm")]
        public static extern bool PlaySound(string szSound, IntPtr hMod, PlaySoundFlags flags);
    }

    public class Sound
    {
        //播放
        public static void Play(string strFileName)
        {
            switch (strFileName)
            {
                case "ok":
                    strFileName =System.AppDomain.CurrentDomain.SetupInformation.ApplicationBase+ "sound/startwin.wav"; 
                    break;
                case "error":
                    strFileName = System.AppDomain.CurrentDomain.SetupInformation.ApplicationBase + "sound/err.wav";
                    break;
                case "rep":
                    strFileName = System.AppDomain.CurrentDomain.SetupInformation.ApplicationBase + "sound/fall.WAV";
                    break;
                case "huiqi":
                    strFileName = System.AppDomain.CurrentDomain.SetupInformation.ApplicationBase + "sound/huiqi.WAV";
                    break;
                case "huiqiend":
                    strFileName = System.AppDomain.CurrentDomain.SetupInformation.ApplicationBase + "sound/huiqiend.WAV";
                    break;
                case "jiangjun":
                    strFileName = System.AppDomain.CurrentDomain.SetupInformation.ApplicationBase + "sound/jiangjun.WAV";
                    break;
                case "kill":
                    strFileName = System.AppDomain.CurrentDomain.SetupInformation.ApplicationBase + "sound/kill.WAV";
                    break;
                case "win":
                    strFileName = System.AppDomain.CurrentDomain.SetupInformation.ApplicationBase + "sound/win.WAV";
                    break;
                case "move":
                    strFileName = System.AppDomain.CurrentDomain.SetupInformation.ApplicationBase + "sound/start.WAV";
                    break;
                case "hold":
                    strFileName = System.AppDomain.CurrentDomain.SetupInformation.ApplicationBase + "sound/stop.WAV";
                    break;
                case "no":
                    strFileName = System.AppDomain.CurrentDomain.SetupInformation.ApplicationBase + "sound/no.WAV";
                    break;
                case "popup":
                    strFileName = System.AppDomain.CurrentDomain.SetupInformation.ApplicationBase + "sound/popup.WAV";
                    break;
                case "mayfall":
                    strFileName = System.AppDomain.CurrentDomain.SetupInformation.ApplicationBase + "sound/mayfall.WAV";
                    break;
            }

            //调用PlaySound方法,播放音乐  
            Sound_Play.PlaySound(strFileName, IntPtr.Zero, Sound_Play.PlaySoundFlags.SND_ASYNC);
        }

        //关闭
        public static void Stop()
        {
            Sound_Play.PlaySound(null, IntPtr.Zero, Sound_Play.PlaySoundFlags.SND_ASYNC);
        }
    }
}
