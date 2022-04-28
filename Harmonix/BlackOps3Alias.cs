using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web;

namespace Harmonix
{
    internal class BlackOps3Alias
    {
        /// <summary>
        /// Holds all the contexts with values
        /// </summary>
        public static Dictionary<string, List<string>> contexts = new Dictionary<string, List<string>>();

        /// <summary>
        /// The default Black Ops 3 Alias Header
        /// </summary>
        public static string[] DefaultHeader =
        {
            "Name",
            "Behavior",
            "Storage",
            "FileSpec",
            "FileSpecSustain",
            "FileSpecRelease",
            "Template",
            "Loadspec",
            "Secondary",
            "SustainAlias",
            "ReleaseAlias",
            "Bus",
            "VolumeGroup",
            "DuckGroup",
            "Duck",
            "ReverbSend",
            "CenterSend",
            "VolMin",
            "VolMax",
            "DistMin",
            "DistMaxDry",
            "DistMaxWet",
            "DryMinCurve",
            "DryMaxCurve",
            "WetMinCurve",
            "WetMaxCurve",
            "LimitCount",
            "LimitType",
            "EntityLimitCount",
            "EntityLimitType",
            "PitchMin",
            "PitchMax",
            "PriorityMin",
            "PriorityMax",
            "PriorityThresholdMin",
            "PriorityThresholdMax",
            "AmplitudePriority",
            "PanType",
            "Pan",
            "Futz",
            "Looping",
            "RandomizeType",
            "Probability",
            "StartDelay",
            "EnvelopMin",
            "EnvelopMax",
            "EnvelopPercent",
            "OcclusionLevel",
            "IsBig",
            "DistanceLpf",
            "FluxType",
            "FluxTime",
            "Subtitle",
            "Doppler",
            "ContextType",
            "ContextValue",
            "ContextType1",
            "ContextValue1",
            "ContextType2",
            "ContextValue2",
            "ContextType3",
            "ContextValue3",
            "Timescale",
            "IsMusic",
            "IsCinematic",
            "FadeIn",
            "FadeOut",
            "Pauseable",
            "StopOnEntDeath",
            "Compression",
            "StopOnPlay",
            "DopplerScale",
            "FutzPatch",
            "VoiceLimit",
            "IgnoreMaxDist",
            "NeverPlayTwice",
            "ContinuousPan",
            "FileSource",
            "FileSourceSustain",
            "FileSourceRelease",
            "FileTarget",
            "FileTargetSustain",
            "FileTargetRelease",
            "Platform",
            "Language",
            "OutputDevices",
            "PlatformMask",
            "WiiUMono",
            "StopAlias",
            "DistanceLpfMin",
            "DistanceLpfMax",
            "FacialAnimationName",
            "RestartContextLoops",
            "SilentInCPZ",
            "ContextFailsafe",
            "GPAD",
            "GPADOnly",
            "MuteVoice",
            "MuteMusic",
            "RowSourceFileName",
            "RowSourceShortName",
            "RowSourceLineNumber"
        };

        /// <summary>
        /// Get the ToolTip for the header
        /// </summary>
        /// <param name="header">Name of the header</param>
        /// <returns></returns>
        public static string GetToolTip(dynamic header)
        {
            if (header != null)
            {
                switch (header.ToString())
                {
                    case "Name":
                        return "Name of this sound alias.";
                    case "Storage":
                        return "(loaded) – Will load the sounds into RAM | (streamed) - Will load the sounds from disk.";
                    case "FileSpec":
                        return "The path and filename of the physical wav file for this sound alias.";
                    case "FileSpecSustain":
                        return "If this column is filled out with a looping asset then it will be triggered when the (one shot) asset specified in FileSpec finishes.";
                    case "FileSpecRelease":
                        return "If this column is filled out then this asset will be triggered when the looping asset specified in FileSpecSustain is stopped.";
                    case "Template":
                        return "This field points to the name of a template alias defined in a template CSV";
                    case "Secondary":
                        return "Specify another sound alias here and it will be triggered immediately after the “primary” sound alias.";
                    case "Bus":
                        return "This is the bus that the sound belongs to. For example, in the sound menu options sliders.";
                    case "VolumeGroup":
                        return "This basically allows you to group together like sounds and specify an attenuation scaler to that group.";
                    case "DuckGroup":
                        return "The name of the duck group that this sound alias belongs to.";
                    case "Duck":
                        return "Specify a duck name that will be triggered when this sound alias plays back.";
                    case "ReverbSend":
                        return "Value in dB SPL that the reverberated (wet) portion of the sound will be attenuated by. (0-100)";
                    case "CenterSend":
                        return "Override the center value specified in the pan for this alias if the specified CenterSend is larger";
                    case "VolMin":
                        return "Specify the volume min and max of the sound in dB SPL and the engine will randomize between these levels.";
                    case "VolMax":
                        return "Specify the volume min and max of the sound in dB SPL and the engine will randomize between these levels.";
                    case "DistMin":
                        return "Is the point in distance units where the sound will begin to “falloff” or attenuate. So the sound will be at max volume from the player until it reaches this distance.";
                    case "DistMaxDry":
                        return "Is the point where it will become silent. It will scale between full volume at DistMin to silent at DistmaxDry.";
                    case "DistMaxWet":
                        return "Is provided to give similar behavior but for the “wet” or reverberated portion of the sound. This allows you to extend the distance in which you will hear the sound reverberating.";
                    case "DryMinCurve":
                        return "This is the curve type that specifies how the sound will falloff depending on distance.";
                    case "DryMaxCurve":
                        return "This is the curve type that specifies how the sound will falloff depending on distance.";
                    case "WetMinCurve":
                        return "This is the curve type that specifies how the sound will falloff depending on distance.";
                    case "WetMaxCurve":
                        return "This is the curve type that specifies how the sound will falloff depending on distance.";
                    case "LimitCount":
                        return "The number of simultaneous sounds of this alias that will play back before limiting will occur.";
                    case "LimitType":
                        return "Once the engine hits the LimitCount for this particular alias a decision needs to be made depending on this type specified.";
                    case "EntityLimitCount":
                        return "Same as LimitCount except rather than taking all playing sounds into account it will limit the instances of this sound alias playing back on a particular entity.";
                    case "EntityLimitType":
                        return "Same as LimitType except specifies the limit type to be applied to a particular entity.";
                    case "PitchMin":
                        return "Specify a min/max pitch in cents for this sound instance to play back at. It will randomize between min/max pitch. +1200 = double pitch, -1200 = half pitch.";
                    case "PitchMax":
                        return "Specify a min/max pitch in cents for this sound instance to play back at. It will randomize between min/max pitch. +1200 = double pitch, -1200 = half pitch.";
                    case "PanType":
                        return "Specifies how the sound changes depending on world position";
                    case "Pan":
                        return "Determines how the sound will be panned throughout the 3D sound field.";
                    case "Futz":
                        return "The name of a futz patch that will be applied to this sound in real-time.";
                    case "Looping":
                        return "Specifies how the sound changes depending on world position";
                    case "RandomizeType":
                        return "If a value is specified here the engine will compute the same random value each time this sound is called on the same entity. If you want actual randomness per entity this field should be left blank.";
                    case "Probability":
                        return "A specified percentage of the sound playing back or not after it’s triggered. 1.0 = 100%, 0.0 – 0% (no playback).";
                    case "StartDelay":
                        return "Value in milliseconds to delay the start of the sound once it is triggered.";
                    case "EnvelopMin":
                        return "When the listener (player) is within this distance from a 3D sound source the sound will be panned into the speakers that would not normally have any of the signal. The amount of signal that is panned into these speakers is determined by the EnvelopPercent.";
                    case "EnvelopMax":
                        return "This distance defines an outer ring around the sound source where the amount of EnvelopPercent is then scaled linearly between the full value and 0 as the listener moves from the min to the max ring.";
                    case "EnvelopPercent":
                        return "The amount of the signal (in dB SPL) that is panned into the channels to produce the envelop effect.";
                    case "OcclusionLevel":
                        return "Float value between 0-1.0 that specifies how much the real-time occlusion level will be scaled.";
                    case "IsBig":
                        return "If set to true this turns off geometry based occlusion when you are close to a 3D sound.";
                    case "DistanceLpf":
                        return "Whether distance based low passed filtering will be applied to this sound. Engine will take the lesser of the calculated distance based LPF value and the occlusion based LPF value.";
                    case "FluxType":
                        return "Flux is a system in the sound engine that allows a sound designer to get 3D positional effects on a sound without having it tied to an actual 3D object in the world.";
                    case "FluxTime":
                        return "Time in milliseconds that the flux effect will be applied for as it travels to DistMaxWet.";
                    case "Subtitle":
                        return "Any subtitle that should be displayed on screen, usually accompanying a dialog line.";
                    case "Doppler":
                        return "Whether this sound will have the Doppler effect applied.";
                    case "ContextType":
                        return "The context type, used for checking if you're indoor / outdoor / underwater and more.";
                    case "ContextValue":
                        return "The value of the context";
                    case "ContextType1":
                        return "The context type, used for checking if you're indoor / outdoor / underwater and more.";
                    case "ContextValue1":
                        return "The value of the context";
                    case "ContextType2":
                        return "The context type, used for checking if you're indoor / outdoor / underwater and more.";
                    case "ContextValue2":
                        return "The value of the context";
                    case "ContextType3":
                        return "The context type, used for checking if you're indoor / outdoor / underwater and more.";
                    case "ContextValue3":
                        return "The value of the context";
                    case "Timescale":
                        return "Whether the sound should be effected by the global timescale setting.";
                    case "IsMusic":
                        return "Depending if the alias is part of the music system.";
                    case "IsCinematic":
                        return "Depending if the alias is part of a cinematic sequence.";
                    case "FadeIn":
                        return "Time in milliseconds that the sound will fade in from silence to full volume.";
                    case "FadeOut":
                        return "Time in milliseconds that the sound will fade out over when it ends.";
                    case "Pauseable":
                        return "Depending on if the sounds should pause when the game pauses.";
                    case "StopOnEntDeath":
                        return "Determine if this sound should respond to a call to stop sounds on an entity.";
                    case "Compression":
                        return "Consoles only as PC is FLAC encoded at a constant compression ratio.";
                    case "StopOnPlay":
                        return "Not currently implemented.";
                    case "DopplerScale":
                        return "Value between -100.0f – 100.0f that scales the Doppler effect on a sound.";
                    case "FutzPatch":
                        return "Name of a futz patch.";
                    case "VoiceLimit":
                        return "Unused.";
                    case "IgnoreMaxDist":
                        return "By default the engine will cull a sound play event if a 3D sound is triggered on an object that is already past the maximum distance specified (DistMaxWet).";
                    case "NeverPlayTwice":
                        return "Set to yes to prevent a sound alias from triggering more than once.";
                    case "oldest":
                        return "Oldest sound of this alias will be stopped immediately.";
                    case "reject":
                        return "This sound alias will be rejected. All others continue.";
                    case "priority":
                        return "The sound instance to be stopped is chosen depending on priority.";
                    case "2d":
                        return "Constant volume/pan. Plays full volume and ignores distance. Music, UI sounds etc.";
                    case "3d":
                        return "Changes volume/pan depending on world position.";
                    case "2.5":
                        return "Ignores distance but pans sound around the player depending on world position.";
                    case "volume":
                        return "Same random volume if sound called on same entity.";
                    case "pitch":
                        return "Same random pitch if sound called on same entity.";
                    case "variant":
                        return "Same random alias variant selection of called on same entity.";
                    case "left_player":
                        return "To the left of the player.";
                    case "right_player":
                        return "To the right of the player.";
                    case "center_player":
                        return "Straight out from the center of the player.";
                    case "random_player":
                        return "Random direction from the player.";
                    case "left_shot":
                        return "To the left of a shot.";
                    case "center_shot":
                        return "Straight out from the center of the shooter.";
                    case "right_shot":
                        return "To the right of a shot.";
                    case "random_direction":
                        return "Completely random in 3D space.";
                }
            }

            return string.Empty;
        }

        /// <summary>
        /// Checks which cells needs Data Validation
        /// </summary>
        /// <param name="header"></param>
        /// <param name="cell"></param>
        public static void CellDataValidation(string header, Range cell)
        {
            // So we can add an empty value
            string emptySpace = HttpUtility.HtmlDecode("&#8206;");

            string BlackOps3Root = Settings.GetBlackOps3Root();

            switch (header)
            {
                case "Storage":
                    SetDataValidation(cell, $"{emptySpace},streamed,loaded");
                    break;
                case "Template":
                    var templateList = new List<string>();

                    var templatePath = Path.Combine(BlackOps3Root, "share", "raw", "sound", "templates");

                    templateList.Add(emptySpace);
                    if (Directory.Exists(templatePath))
                    {
                        var files = Directory.GetFiles(templatePath, "*.csv*");
                        foreach (var file in files)
                        {
                            ReadCSV(templateList, file);
                        }
                    }

                    SetDataValidation(cell, string.Join(",", templateList.ToArray()));
                    break;
                case "Bus":
                    SetDataValidation(cell, $"{emptySpace},BUS_FX,BUS_VOICE,BUS_PFUTZ,BUS_HDRFX,BUS_UI,BUS_MUSIC,BUS_MOVIE,BUS_REFERENCE");
                    break;
                case "VolumeGroup":
                    var volumeGroupList = new List<string>();

                    volumeGroupList.Add(emptySpace);

                    var volumeGroup = Path.Combine(BlackOps3Root, "share", "raw", "sound", "globals", "volume_group.csv");

                    ReadCSV(volumeGroupList, volumeGroup);

                    SetDataValidation(cell, string.Join(",", volumeGroupList.ToArray()));
                    break;
                case "DuckGroup":
                    var duckGroupList = new List<string>();

                    duckGroupList.Add(emptySpace);

                    var duckGroup = Path.Combine(BlackOps3Root, "share", "raw", "sound", "globals", "duck_group.csv");

                    ReadCSV(duckGroupList, duckGroup);

                    SetDataValidation(cell, string.Join(",", duckGroupList.ToArray()));
                    break;
                case "Duck":
                    var ducks = new List<string>();

                    var duckPath = Path.Combine(BlackOps3Root, "share", "raw", "sound", "ducks");

                    ducks.Add(emptySpace);

                    if (Directory.Exists(duckPath))
                    {
                        var files = Directory.GetFiles(duckPath, "*.duk*");

                        foreach (var file in files)
                        {
                            ducks.Add(Path.GetFileNameWithoutExtension(file));
                        }
                    }

                    SetDataValidation(cell, string.Join(",", ducks.ToArray()));
                    break;
                case "DryMinCurve":
                case "DryMaxCurve":
                case "WetMinCurve":
                case "WetMaxCurve":
                    var curveList = new List<string>();

                    curveList.Add(emptySpace);

                    var curve = Path.Combine(BlackOps3Root, "share", "raw", "sound", "globals", "curve.csv");

                    ReadCSV(curveList, curve);

                    SetDataValidation(cell, string.Join(",", curveList.ToArray()));
                    break;
                case "LimitType":
                case "EntityLimitType":
                    SetDataValidation(cell, $"{emptySpace},none,oldest,reject,priority");
                    break;
                case "PanType":
                    SetDataValidation(cell, $"{emptySpace},2d,3d,2.5");
                    break;
                case "Pan":
                    var panList = new List<string>();

                    panList.Add(emptySpace);

                    var pan = Path.Combine(BlackOps3Root, "share", "raw", "sound", "globals", "pan.csv");

                    ReadCSV(panList, pan);

                    SetDataValidation(cell, string.Join(",", panList.ToArray()));
                    break;
                case "Futz":
                case "FutzPatch":
                    var futzList = new List<string>();

                    futzList.Add(emptySpace);

                    var futz = Path.Combine(BlackOps3Root, "share", "raw", "sound", "globals", "futz.csv");

                    ReadCSV(futzList, futz);

                    SetDataValidation(cell, string.Join(",", futzList.ToArray()));
                    break;
                case "Looping":
                    SetDataValidation(cell, $"{emptySpace},nonlooping,looping");
                    break;
                case "RandomizeType":
                    SetDataValidation(cell, $"{emptySpace},volume,pitch,variant");
                    break;
                case "IsBig":
                case "DistanceLpf":
                case "Doppler":
                case "Timescale":
                case "IsMusic":
                case "IsCinematic":
                case "Pauseable":
                case "StopOnEntDeath":
                case "IgnoreMaxDist":
                case "NeverPlayTwice":
                    SetDataValidation(cell, $"{emptySpace},yes,no");
                    break;
                case "FluxType":
                    SetDataValidation(cell, $"{emptySpace},left_player,right_player,center_player,random_player,left_shot,center_shot,right_shot,random_direction");
                    break;
                case "ContextType":
                case "ContextType1":
                case "ContextType2":
                case "ContextType3":
                    if (contexts.Count <= 0)
                    {
                        var contextPath = Path.Combine(BlackOps3Root, "share", "raw", "sound", "contexts");

                        contexts.Add(emptySpace, new List<string>());

                        if (Directory.Exists(contextPath))
                        {
                            var files = Directory.GetFiles(contextPath, "*.ctx*");

                            foreach (var file in files)
                            {
                                var data = File.ReadAllText(file);

                                BlackOps3Context context = JsonConvert.DeserializeObject<BlackOps3Context>(data);

                                contexts.Add(Path.GetFileNameWithoutExtension(file), context.CtxValues);
                            }
                        }
                    }

                    var keys = contexts.Keys.ToList();
                    SetDataValidation(cell, string.Join(",", keys.ToArray()));
                    break;
            }
        }

        /// <summary>
        /// Set the Data Validation on a cell
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="list"></param>
        public static void SetDataValidation(Range cell, string list)
        {
            cell.Validation.Delete();
            cell.Validation.Add(XlDVType.xlValidateList, XlDVAlertStyle.xlValidAlertInformation, XlFormatConditionOperator.xlBetween, list, Type.Missing);

            cell.Validation.IgnoreBlank = false;
            cell.Validation.InCellDropdown = true;
        }

        /// <summary>
        /// Data validation on a context cell
        /// </summary>
        /// <param name="Target"></param>
        /// <param name="Initial"></param>
        public static void DataValidationContext(Range Target, bool Initial = false)
        {
            Range header = Harmonix.GetActiveWorksheet().Cells[1, Target.Column];

            if (header.Value != null)
            {
                string headerName = Harmonix.GetActiveWorksheet().Cells[1, Target.Column].Value.ToString();

                switch (headerName)
                {
                    case "ContextType":
                    case "ContextType1":
                    case "ContextType2":
                    case "ContextType3":
                        if (contexts.Count > 0)
                        {
                            Range cell = Harmonix.GetActiveWorksheet().Cells[Target.Row, Target.Column + 1];

                            if (cell != null)
                            {
                                cell.Validation.Delete();
                                if(!Initial)
                                {
                                    cell.Value = "";
                                }
                                
                            }

                            if (!string.IsNullOrWhiteSpace(Target.Value))
                            {
                                string value = Target.Value.ToString();

                                if (contexts.ContainsKey(value))
                                {
                                    List<string> values = contexts[value];

                                    if (values.Count > 0)
                                    {
                                        SetDataValidation(cell, string.Join(",", values.ToArray()));
                                        if(!Initial)
                                        {
                                            cell.Value = values[0];
                                        }
                                        
                                    }
                                }
                            }
                        }
                        break;
                }
            }
        }

        /// <summary>
        /// Read a CSV file
        /// </summary>
        /// <param name="list"></param>
        /// <param name="file"></param>
        public static void ReadCSV(List<string> list, string file)
        {
            using (StreamReader sr = new StreamReader(file))
            {
                // Skip first row
                sr.ReadLine();
                while (sr.Peek() >= 0)
                {
                    string line = sr.ReadLine();

                    if (line != null)
                    {
                        string template = line.Split(',')[0].Trim();

                        if (!string.IsNullOrWhiteSpace(template) && !template.StartsWith("#"))
                        {
                            list.Add(template);
                        }
                    }
                }
            }
        }
    }
}
