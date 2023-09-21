# HeyPowerPoint MSHackathon2023

https://hackbox.microsoft.com/hackathons/hackathon2023/project/45877

Windows
1. [Grab Visual Studio](https://visualstudio.microsoft.com/)
2. [Use an Azure subscription and create a speech resource in Azure](https://learn.microsoft.com/en-us/azure/ai-services/speech-service/get-started-speech-to-text?tabs=windows%2Cterminal&pivots=programming-language-csharp#prerequisites)
3. Replace `subscriptionKey` and `region` with the azure speech service's key and region.
4. Open up a presentation in PowerPoint
5. Run the HeyPowerPoint project
6. Click on `Start Listening`

--------

Default transition triggers:
- Start - Connects to an open PowerPoint presentation and begins the SlideShow
- Next - Transition to the next slide
- Back - Transition to the previous slide numerically
- End - End the slide show presentation
- [number] - Transition to slide # [number] if its within bounds

--------

Goal

'Hey PowerPoint!' aims to provide Microsoft PowerPoint users a customizable slide transition tool featuring a state-machine powered with Azure AI's Speech service to recognize user-preset phrases to trigger animations, segue from slide i -> i+1, or callback to a previous slide/sneak peek to a later slide.
