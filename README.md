# vsto-outlook-busylight

There is a company here in Denmark that makes a ["busy light"](http://www.plenom.com/products/kuando-busylight-uc-for-skype4b-lync-cisco-jabber-more/) to give others around your office an idea of whether they should approach you or not. Not only is it a light, but it has built-in speakers! It can can sit on top of your monitor or attach to the wall outside your home office and light up based on a status you set manually or based on a status it gets seamlessly through Skype.

That's great and all, especially because we work in an open office and we each get our own light, but surely we can come up with something a bit more creative. For example, I wanted to set mine up to react when my code check-ins pass or fail. I want a quick alert for myself, and I hope that the public nature of the alerts will encourage me to get more green lights than red.

Fortunately, Plenom, the maker of the light, has an SDK to make controlling the light and sound of the device. The bad news, though, is that it's not so easy to see what each method does. What is the difference between "Alert" and "AlertAndReturn"? Spoiler: The latter does the same as the first, but ignores the color you choose and flashes blue instead. Odd, I know.

So I made two seperate projects. The first is the tool I wish someone had made before I started - a simple console app which would let me see the difference between alert and jingle, hear the different sounds, and see just how loud I can ring the BusyLight before I annoy my colleagues. There isn't much to explain, so check out the source (and an exe to try it yourself) on [github](https://github.com/hoovercj/BusyLight-Demo-CSharp).

I get email alerts about the status of my check-ins, so the second project is a [VSTO Add-In](https://msdn.microsoft.com/en-us/library/ms268878.aspx) for Outlook. VSTO stands for "Visual Studio Tools for Office" and is a type of .NET add-in that installs on your system and boots with your office product. This is different than the newer ["Office Add-Ins Platform"](http://dev.office.com/docs/add-ins/overview/office-add-ins) built with javascript/html/css. The add-in is very simple and is based on [this msdn article](https://msdn.microsoft.com/en-us/library/cc668191.aspx).

The main steps are:

* Create a Visual Studio project from a VSTO template.
* Call your code from the events or hooks that the template exposes in `ThisAddIn.cs`. A simple version of the code needed for my add-in is below.
* Create and use a [Click-Once installer](https://msdn.microsoft.com/en-us/library/bb772100.aspx) for the add-in
* Annoy your teammates with a singing success-light! Or embarrass yourself with a singing shame-light...

```csharp
// Template method filled with my code
void items_ItemAdd(object Item)
{
    if (Item == null) { return; }
    Outlook.MailItem mail = (Outlook.MailItem)Item;

    if (mail.MessageClass != "IPM.Note" || !mail.Sender.Address.ToLower().Equals("example@email.com") { return; }

    var sdk = new SDK(); // BusyLight SDK

    if (mail.Subject.ToLower().Contains("completed successfully"))
    {
        sdk.Alert(BusylightColor.Green, BusylightSoundClip.FairyTale, BusylightVolume.Low);
    }
    else if (mail.Subject.ToLower().Contains("failed"))
    {
        sdk.Alert(BusylightColor.Red, BusylightSoundClip.Funky, BusylightVolume.Low);
    }
    else { return; }
    System.Threading.Thread.Sleep(3000);
    sdk.Light(BusylightColor.Off);
}
```

Feel free to leave a comment in the [issues](https://github.com/hoovercj/vsto-outlook-busylight/issues) session if you have any questions.

