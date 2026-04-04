namespace MS.Models;
public class tbl_post
{
    public int? id { get; set; }
    public int? PatientIE_ID { get; set; }
    public int? PatientFU_ID { get; set; }

    public string? CreationDate { get; set; }
    public string? Patient_Name { get; set; }

    public string? DOB { get; set; }
    public string? DOS { get; set; }

    public bool? chkLeftShoulder { get; set; }
    public bool? chkRightShoulder { get; set; }
    public bool? chkLeftHip { get; set; }
    public bool? chkRightHip { get; set; }
    public bool? chkLeftKnee { get; set; }
    public bool? chkRightKnee { get; set; }
    public bool? chkLeftAnkleFoot { get; set; }
    public bool? chkRightAnkleFoot { get; set; }

    public string? txtHistoryPresentillness { get; set; }   
    public string? txtPhysicalExamination { get; set; }   
    public string? txtExaminedResult { get; set; }   


}


