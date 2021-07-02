

#Totally works great!
#Charlie's laptop is CHL001405
invoke-command -ComputerName "chl001405" -ScriptBlock {
    
    
    
    Add-Type -AssemblyName System.speech;
    $speak = New-Object System.Speech.Synthesis.SpeechSynthesizer;
    
    $msg = "Thank you for subscribing to CatFacts!";
    $msg2 = "The cat is a small carnivorous mammal. It is the only domesticated species in the family Felidae and often referred to as the domestic cat to distinguish it from wild members of the family";
    $msg3 = "The cat is either a house cat or a farm cat, which are pets, or a feral cat, which ranges freely and avoids human contact"
    $msg4 = "A house cat is valued by humans for companionship and for its ability to hunt rodents.";
    $msg5 = "About 60 cat breeds are recognized by various cat registries.";
    $msg6 = "Cats are similar in anatomy to the other felid species, with a strong flexible body, quick reflexes, sharp teeth and retractable claws adapted to killing small prey. They are predators who are most active at dawn and dusk (crepuscular).";
    $msg7 = "Cats can hear sounds too faint or too high in frequency for human ears, such as those made by mice and other small animals.";
    $msg8 = "Compared to humans, they see better in the dark (they see in near total darkness) and have a better sense of smell, but poorer color vision.";
    $speak.Speak("$msg");
    $speak.Speak("$msg2");
    $speak.Speak("$msg3");
    $speak.Speak("$msg4");
    $speak.Speak("$msg5");
    $speak.Speak("$msg6");
    $speak.Speak("$msg7");
    $speak.Speak("$msg8")

}

$computers = "chl001405"
foreach ($computer in $computers)
{
#$name = read-host "Enter computer name "
$msg = "Thank you for subscribing to CatFacts!"
$msg2 = "The cat is a small carnivorous mammal. It is the only domesticated species in the family Felidae and often referred to as the domestic cat to distinguish it from wild members of the family"
$msg3 = "The cat is either a house cat or a farm cat, which are pets, or a feral cat, which ranges freely and avoids human contact"
$msg4 = "A house cat is valued by humans for companionship and for its ability to hunt rodents."
$msg5 = "About 60 cat breeds are recognized by various cat registries."
$msg6 = "Cats are similar in anatomy to the other felid species, with a strong flexible body, quick reflexes, sharp teeth and retractable claws adapted to killing small prey. They are predators who are most active at dawn and dusk (crepuscular)."
$msg7 = "Cats can hear sounds too faint or too high in frequency for human ears, such as those made by mice and other small animals."
$msg8 = "Compared to humans, they see better in the dark (they see in near total darkness) and have a better sense of smell, but poorer color vision."

#$speak.Speak("$msg")
Invoke-WmiMethod -Path Win32_Process -Name Create -ArgumentList "msg * $msg" -ComputerName $computer
Invoke-WmiMethod -Path Win32_Process -Name Create -ArgumentList "msg * $msg2" -ComputerName $computer
Invoke-WmiMethod -Path Win32_Process -Name Create -ArgumentList "msg * $msg3" -ComputerName $computer
Invoke-WmiMethod -Path Win32_Process -Name Create -ArgumentList "msg * $msg4" -ComputerName $computer
Invoke-WmiMethod -Path Win32_Process -Name Create -ArgumentList "msg * $msg5" -ComputerName $computer
Invoke-WmiMethod -Path Win32_Process -Name Create -ArgumentList "msg * $msg6" -ComputerName $computer
Invoke-WmiMethod -Path Win32_Process -Name Create -ArgumentList "msg * $msg7" -ComputerName $computer
Invoke-WmiMethod -Path Win32_Process -Name Create -ArgumentList "msg * $msg8" -ComputerName $computer
}
