using System;
using System.Collections.Generic;
using System.Text;
 
using System.Runtime.InteropServices;
using System.Threading;
using System.Diagnostics;

namespace LyncFellow
{
    class Buddies
    {
        enum VendorID { TenxTechnology = 0x1130 }
        enum TenxTechDeviceID { iBuddy1 = 0x0001, iBuddy2 = 0x0002 }

        Dictionary<string, iBuddy> _buddies = new Dictionary<string,iBuddy>();
        public int LastWin32Error = 0;

        public string currentUser = "self";
        public string currentUserUri = "";
        public string currentUserName = "";
        public string currentUserActivity = "";
        public object ActivityString;
        public object DisplayName;
        public object PersonalNote;
        public object Desc;
        public object Availability;
        public List<string> users = new List<string>();
        public List<string> users_group = new List<string>();

        private iBuddy.Color _color = iBuddy.Color.Off;

        public bool thisDebug = false;          //debug log

        public void RefreshList()
        {
            LastWin32Error = 0;

            var keys = new string[_buddies.Keys.Count];
            _buddies.Keys.CopyTo(keys, 0);
            foreach (var Key in keys)
            {
                var buddy = _buddies[Key];
                if (buddy.LastWin32Error != 0)
                {
                    LastWin32Error = buddy.LastWin32Error;
                }
                if (!buddy.IsAlive)
                {
                    buddy.Release();
                    _buddies.Remove(Key);
                }
            }

            List<string> devicePaths = new List<string>();
            devicePaths.AddRange(HIDDevices.Find((int)VendorID.TenxTechnology, (int)TenxTechDeviceID.iBuddy1));
            devicePaths.AddRange(HIDDevices.Find((int)VendorID.TenxTechnology, (int)TenxTechDeviceID.iBuddy2));
            foreach(var devicePath in devicePaths)
            {
                if (!_buddies.ContainsKey(devicePath))
                {
                    var buddy = new iBuddy();
                    if (buddy.Init(devicePath))
                    {
                        _buddies.Add(devicePath, buddy);
                    }
                }
            }

        }

        public int Count { get { return _buddies.Count; } }

        public void Release()
        {
            foreach (var buddy in _buddies.Values)
            {
                buddy.Release();
            }
            _buddies.Clear();
        }

        public iBuddy.Color Color
        {
            get
            { 
                return _color;
            }
            set
            {
                _color = value;
                UpdateColor();
            }
        }

        public void UpdateColor()
        {
            if (thisDebug)
            {
                Trace.WriteLine("UPDATE STATE Buddies");
            }
            foreach (var buddy in _buddies.Values)
            {
                buddy.SetColor(_color);
            }
        }

        public void Dance(uint TimeMs = 5000)
        {
            foreach (var buddy in _buddies.Values)
            {
                buddy.Dance(TimeMs);
            }
        }

        public void FlapWings(uint TimeMs = 5000)
        {
            foreach (var buddy in _buddies.Values)
            {
                buddy.FlapWings(TimeMs);
            }
        }

        public void Heartbeat(uint TimeMs = 5000)
        {
            foreach (var buddy in _buddies.Values)
            {
                buddy.Heartbeat(TimeMs);
            }
        }

        public void Rainbow(uint TimeMs = 3000)
        {
            foreach (var buddy in _buddies.Values)
            {
                buddy.Rainbow(TimeMs);
            }
        }

        public void FreakOut(uint TimeMs = 3000)
        {
            Dance(1200);

            foreach (var buddy in _buddies.Values)
            {
                buddy.FlapWings(TimeMs);
                buddy.Rainbow(TimeMs);
                buddy.Heartbeat(TimeMs);
            }

 

        }

    }
}
