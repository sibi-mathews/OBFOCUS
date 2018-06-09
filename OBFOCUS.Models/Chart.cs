using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OBFOCUS.Models
{
    class Chart
    {
        private int _id = 0;
        private int _patientID = 0;
        private string _medicalRecord = string.Empty;
        private string _lastName = string.Empty;
        private string _firstName = string.Empty;
        private DateTime _dob = DateTime.MinValue;
        private string _race = string.Empty;
        private int _gravida = 0;
        private int _para = 0;
        private string _sab = string.Empty;
        private string _top = string.Empty;
        private string _term = string.Empty;
        private string _living = string.Empty;
        private DateTime _edc = DateTime.MinValue;
        private DateTime _lmp = DateTime.MinValue;
        private DateTime _earlyUS = DateTime.MinValue;
        private string _useEDCBy = string.Empty;
        private int _refDX = 0;
        private int _physicianID = 0;
        private int _delHospitalID = 0;
        private int _siteID = 0;
        private DateTime _dateCreated = DateTime.MinValue;
        private string _pLastName = string.Empty;
        private string _type = string.Empty;
        private string _rh = string.Empty;
        private string _antiBody = string.Empty;
        private int _preWeight = 0;
        private int _height = 0;
        private short _tab = 0;
        private DateTime _examDate = DateTime.MinValue;
        private int _examID = 0;
        private string _signed = string.Empty;
        private string _socialSecurity = string.Empty;
        private short _normal = 0;
        private string _normalComments = string.Empty;
        private short _bleeding = 0;
        private string _bleedingComments = string.Empty;
        private short _cramping = 0;
        private string _crampingComments = string.Empty;
        private short _excess = 0;
        private string _excessComments = string.Empty;
        private short _radiation = 0;
        private string _radiationComments = string.Empty;
        private short _chemicals = 0;
        private string _chemicalsComments = string.Empty;
        private short _smoking = 0;
        private string _smokingComments = string.Empty;
        private short _alcohol = 0;
        private string _alcoholComments = string.Empty;
        private short _drugs = 0;
        private string _drugsComments = string.Empty;
        private short _fever = 0;
        private string _feverComments = string.Empty;
        private short _medicalHx = 0;
        private string _medicalHistory = string.Empty;
        private short _surgicalHx = 0;
        private string _surgicalHistory = string.Empty;
        private short _gynHx = 0;
        private string _gynHistory = string.Empty;
        private short _familyHx = 0;
        private string _familyHistory = string.Empty;
        private string _socialHistory = string.Empty;
        private string _transfusion = string.Empty;
        private short _birthDefectsMat = 0;
        private string _defectsMatco = string.Empty;
        private short _birthDefectsPat = 0;
        private string _defectsPatco = string.Empty;
        private string _allergies = string.Empty;
        private string _billingAddress1 = string.Empty;
        private short _patientAutoNum = 0;
        private int _examNumber = 0;
        private string _medications = string.Empty;
        private string _userID = string.Empty;
        private DateTime _updatedDate = DateTime.MinValue;
        private string _updatedBy = string.Empty;
        private DateTime _lastOpenedDate = DateTime.MinValue;
        private string _lastOpenedBy = string.Empty;
        private int _tsUpdate = 0;
        private string _workstation = string.Empty;
        private short _bReadOnly = 0;
        private int _defaultExaminerID = 0;
        private bool _isLab = false;
        private bool _documentReviewed = false;
        private string _reviewComments = string.Empty;
        private string _labSiteDescrip = string.Empty;
        private bool _faxOptIn = false;
        private bool _emailOptIn = false;
        private bool _mailOptIn = false;
        private bool _otherOptIn = false;
        private string _roomNumber = string.Empty;
        private string _docRecLab = string.Empty;
        private string _docRecStat = string.Empty;
        private int _docRecExamID = 0;
        private string _docRecExamDate = string.Empty;
        private string _documentPath = string.Empty;

        public int Id
        {
            get { return _id; }
            set { _id = value; }
        }

        public int PatientId
        {
            get { return _patientID; }
            set { _patientID = value; }
        }

        public string MedicalRecord
        {
            get { return _medicalRecord; }
            set { _medicalRecord = value; }
        }

        public string LastName
        {
            get { return _lastName; }
            set { _lastName = value; }
        }

        public string FirstName
        {
            get { return _firstName; }
            set { _firstName = value; }
        }

        public DateTime DOB
        {
            get { return _dob; }
            set { _dob = value; }
        }

        public string Race
        {
            get { return _race; }
            set { _race = value; }
        }

        public int Gravida
        {
            get { return _gravida; }
            set { _gravida = value; }
        }

        public int Para
        {
            get { return _para; }
            set { _para = value; }
        }

        public string Sab
        {
            get { return _sab; }
            set { _sab = value; }
        }

        public string Top
        {
            get { return _top; }
            set { _top = value; }
        }

        public string Term
        {
            get { return _term; }
            set { _term = value; }
        }

        public string Living
        {
            get { return _living; }
            set { _living = value; }
        }

        public DateTime EDC
        {
            get { return _edc; }
            set { _edc = value; }
        }

        public DateTime LMP
        {
            get { return _lmp; }
            set { _lmp = value; }
        }

        public DateTime EarlyUS
        {
            get { return _earlyUS; }
            set { _earlyUS = value; }
        }

        public string UseEDCBy
        {
            get { return _useEDCBy; }
            set { _useEDCBy = value; }
        }

        public int RefDX
        {
            get { return _refDX; }
            set { _refDX = value; }
        }

        public int PhysicianID
        {
            get { return _physicianID; }
            set { _physicianID = value; }
        }

        public int DelHospitalID
        {
            get { return _delHospitalID; }
            set { _delHospitalID = value; }
        }

        public int SiteID
        {
            get { return _siteID; }
            set { _siteID = value; }
        }

        public DateTime DateCreated
        {
            get { return _dateCreated; }
            set { _dateCreated = value; }
        }

        public string PLastName
        {
            get { return _pLastName; }
            set { _pLastName = value; }
        }

        public string Type
        {
            get { return _type; }
            set { _type = value; }
        }

        public string RH
        {
            get { return _rh; }
            set { _rh = value; }
        }

        public string AntiBody
        {
            get { return _antiBody; }
            set { _antiBody = value; }
        }

        public int PreWeight
        {
            get { return _preWeight; }
            set { _preWeight = value; }
        }

        public int Height
        {
            get { return _height; }
            set { _height = value; }
        }

        public short Tab
        {
            get { return _tab; }
            set { _tab = value; }
        }

        public DateTime ExamDate
        {
            get { return _examDate; }
            set { _examDate = value; }
        }

        public int ExamID
        {
            get { return _examID; }
            set { _examID = value; }
        }

        public string Signed
        {
            get { return _signed; }
            set { _signed = value; }
        }

        public string SocialSecurity
        {
            get { return _socialSecurity; }
            set { _socialSecurity = value; }
        }

        public short Normal
        {
            get { return _normal; }
            set { _normal = value; }
        }

        public string NormalComments
        {
            get { return _normalComments; }
            set { _normalComments = value; }
        }

        public short Bleeding
        {
            get { return _bleeding; }
            set { _bleeding = value; }
        }

        public string BleedingComments
        {
            get { return _bleedingComments; }
            set { _bleedingComments = value; }
        }

        public short Cramping
        {
            get { return _cramping; }
            set { _cramping = value; }
        }

        public string CrampingComments
        {
            get { return _crampingComments; }
            set { _crampingComments = value; }
        }

        public short Excess
        {
            get { return _excess; }
            set { _excess = value; }
        }

        public string ExcessComments
        {
            get { return _excessComments; }
            set { _excessComments = value; }
        }


        public short Radiation
        {
            get { return _radiation; }
            set { _radiation = value; }
        }

        public string RadiationComments
        {
            get { return _radiationComments; }
            set { _radiationComments = value; }
        }


        public short Chemicals
        {
            get { return _chemicals; }
            set { _chemicals = value; }
        }

        public string ChemicalsComments
        {
            get { return _chemicalsComments; }
            set { _chemicalsComments = value; }
        }


        public short Smoking
        {
            get { return _smoking; }
            set { _smoking = value; }
        }

        public string SmokingComments
        {
            get { return _smokingComments; }
            set { _smokingComments = value; }
        }


        public short Alcohol
        {
            get { return _alcohol; }
            set { _alcohol = value; }
        }

        public string AlcoholComments
        {
            get { return _alcoholComments; }
            set { _alcoholComments = value; }
        }


        public short Drugs
        {
            get { return _drugs; }
            set { _drugs = value; }
        }

        public string DrugsComments
        {
            get { return _drugsComments; }
            set { _drugsComments = value; }
        }


        public short Fever
        {
            get { return _fever; }
            set { _fever = value; }
        }

        public string FeverComments
        {
            get { return _feverComments; }
            set { _feverComments = value; }
        }

        public short MedicalHx
        {
            get { return _medicalHx; }
            set { _medicalHx = value; }
        }

        public string MedicalHistory
        {
            get { return _medicalHistory; }
            set { _medicalHistory = value; }
        }

        public short SurgicalHx
        {
            get { return _surgicalHx; }
            set { _surgicalHx = value; }
        }

        public string SurgicalHistory
        {
            get { return _surgicalHistory; }
            set { _surgicalHistory = value; }
        }

        public short GynHx
        {
            get { return _gynHx; }
            set { _gynHx = value; }
        }

        public string GynHistory
        {
            get { return _gynHistory; }
            set { _gynHistory = value; }
        }

        public short FamilyHx
        {
            get { return _familyHx; }
            set { _familyHx = value; }
        }

        public string FamilyHistory
        {
            get { return _familyHistory; }
            set { _familyHistory = value; }
        }

        public string SocialHistory
        {
            get { return _socialHistory; }
            set { _socialHistory = value; }
        }

        public string Transfusion
        {
            get { return _transfusion; }
            set { _transfusion = value; }
        }

        public short BirthDefectsMat
        {
            get { return _birthDefectsMat; }
            set { _birthDefectsMat = value; }
        }

        public string DefectsMatco
        {
            get { return _defectsMatco; }
            set { _defectsMatco = value; }
        }

        public short BirthDefectsPat
        {
            get { return _birthDefectsPat; }
            set { _birthDefectsPat = value; }
        }

        public string DefectsPatco
        {
            get { return _defectsPatco; }
            set { _defectsPatco = value; }
        }

        public string Allergies
        {
            get { return _allergies; }
            set { _allergies = value; }
        }

        public string BillingAddress1
        {
            get { return _billingAddress1; }
            set { _billingAddress1 = value; }
        }

        public short PatientAutoNum
        {
            get { return _patientAutoNum; }
            set { _patientAutoNum = value; }
        }

        public int ExamNumber
        {
            get { return _examNumber; }
            set { _examNumber = value; }
        }

        public string Medications
        {
            get { return _medications; }
            set { _medications = value; }
        }

        public string UserID
        {
            get { return _userID; }
            set { _userID = value; }
        }

        public DateTime UpdatedDate
        {
            get { return _updatedDate; }
            set { _updatedDate = value; }
        }

        public string UpdatedBy
        {
            get { return _updatedBy; }
            set { _updatedBy = value; }
        }

        public DateTime LastOpenedDate
        {
            get { return _lastOpenedDate; }
            set { _lastOpenedDate = value; }
        }

        public string LastOpenedBy
        {
            get { return _lastOpenedBy; }
            set { _lastOpenedBy = value; }
        }


        public int TSUpdate
        {
            get { return _tsUpdate; }
            set { _tsUpdate = value; }
        }


        public string Workstation
        {
            get { return _workstation; }
            set { _workstation = value; }
        }


        public short BReadOnly
        {
            get { return _bReadOnly; }
            set { _bReadOnly = value; }
        }

        public int DefaultExaminerID
        {
            get { return _defaultExaminerID; }
            set { _defaultExaminerID = value; }
        }

        public bool IsLab
        {
            get { return _isLab; }
            set { _isLab = value; }
        }

        public bool DocumentReviewed
        {
            get { return _documentReviewed; }
            set { _documentReviewed = value; }
        }

        public string ReviewComments
        {
            get { return _reviewComments; }
            set { _reviewComments = value; }
        }

        public string LabSiteDescrip
        {
            get { return _labSiteDescrip; }
            set { _labSiteDescrip = value; }
        }

        public bool FaxOptIn
        {
            get { return _faxOptIn; }
            set { _faxOptIn = value; }
        }

        public bool EmailOptIn
        {
            get { return _emailOptIn; }
            set { _emailOptIn = value; }
        }

        public bool MailOptIn
        {
            get { return _mailOptIn; }
            set { _mailOptIn = value; }
        }

        public bool OtherOptIn
        {
            get { return _otherOptIn; }
            set { _otherOptIn = value; }
        }

        public string RoomNumber
        {
            get { return _roomNumber; }
            set { _roomNumber = value; }
        }

        public string DocRecLab
        {
            get { return _docRecLab; }
            set { _docRecLab = value; }
        }


        public string DocRecStat
        {
            get { return _docRecStat; }
            set { _docRecStat = value; }
        }

        public int DocRecExamID
        {
            get { return _docRecExamID; }
            set { _docRecExamID = value; }
        }

        public string DocRecExamDate
        {
            get { return _docRecExamDate; }
            set { _docRecExamDate = value; }
        }

        public string DocumentPath
        {
            get { return _documentPath; }
            set { _documentPath = value; }
        }
    }
}
