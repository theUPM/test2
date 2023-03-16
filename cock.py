import win32com.client

outlook = win32com.client.Dispatch("Outlook.Application")
appt = outlook.CreateItem(1)  # 1 for AppointmentItem

# Set appointment properties
appt.Subject = "Meeting with John"
appt.Start = "2023-03-17 15:30"
appt.Duration = 60
appt.Location = "Conference Room 1"
appt.Body = "Discuss project status"

appt.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/string/{00020329-0000-0000-C000-000000000046}/PidLidReminderFileParameter", "C:\\reminder_sound.wav") # PidLidReminderFileParameter
appt.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/string/{00020329-0000-0000-C000-000000000046}/PidLidReminderOverride", True) # PidLidReminderOverride

# Save the appointment as a .msg file
appt.SaveAs("D:\\appointment.msg", 3)  # 3 for olMSG