using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MeetingRoom
{
    class DateTimeSlot
    {
        string date;
        int from;
        int to;
        string fromString;
        string tooString;
        string meetingRoomSelected;
        List<string> meetingRoomList;

        public string Date { get => date; set => date = value; }
        public int From { get => from; set => from = value; }
        public int To { get => to; set => to = value; }
        public List<string> MeetingRoomList { get => meetingRoomList; set => meetingRoomList = value; }
        public string MeetingRoomSelected { get => meetingRoomSelected; set => meetingRoomSelected = value; }
        public string FromString { get => fromString; set => fromString = value; }
        public string TooString { get => tooString; set => tooString = value; }
    }
}
