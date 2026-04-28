
using System;

string tenantId = "bbdc798e-3108-46e0-8db9-cad38b09423a";
string clientId = "31c75abf-1594-48ae-9966-8cb04c97632f";
string clientSecret = "BRM8Q~LVRMqCH4zRTr8GMX~cI4p2ABvyiBI6ZbCk";

string senderEmail = "ahmed.raza@awesometechinc.com";
string targetEmail = "m.huzaifa14441@gmail.com";
string targetName = "Muhammad Huzaifa";

var emailService = new EmailService(tenantId, clientId, clientSecret);

List<string> attachments = new List<string>
{
    @"C:\Data\StudentsData.xlsx",
    @"C:\Data\Mern.pdf"
};
Console.WriteLine("🚀 Process Start: Sending professional email with attachments...");

await emailService.SendProfessionalEmailAsync(senderEmail, targetEmail, targetName, attachments);

Console.WriteLine("🏁 Process End.");