// ***********************************************************************
// Assembly         : RecordingBot.Services
// 
// Created          : 09-07-2020
//

// Last Modified On : 09-07-2020
// ***********************************************************************
// <copyright file="CallHandler.cs" company="Microsoft">
//     Copyright �  2020
// </copyright>
// <summary></summary>
// ***********************************************************************>

using Microsoft.CognitiveServices.Speech;
using Microsoft.CognitiveServices.Speech.Audio;
using Microsoft.Graph;
using Microsoft.Graph.Auth;
using Microsoft.Graph.Communications.Calls;
using Microsoft.Graph.Communications.Calls.Media;
using Microsoft.Graph.Communications.Common.Telemetry;
using Microsoft.Graph.Communications.Resources;
using Microsoft.Identity.Client;
using RecordingBot.Model.Constants;
using RecordingBot.Services.Contract;
using RecordingBot.Services.ServiceSetup;
using RecordingBot.Services.Util;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.Timers;


namespace RecordingBot.Services.Bot
{
    /// <summary>
    /// Call Handler Logic.
    /// </summary>
    public class CallHandler : HeartbeatHandler
    {
        /// <summary>
        /// Gets the call.
        /// </summary>
        /// <value>The call.</value>
        public ICall Call { get; }

        /// <summary>
        /// Gets the bot media stream.
        /// </summary>
        /// <value>The bot media stream.</value>
        public BotMediaStream BotMediaStream { get; private set; }


        /// <summary>
        /// The settings
        /// </summary>
        private readonly AzureSettings _settings;
        /// <summary>
        /// The event publisher
        /// </summary>
        private readonly IEventPublisher _eventPublisher;

        /// <summary>
        /// The capture
        /// </summary>
        private CaptureEvents _capture;

        /// <summary>
        /// The is disposed
        /// </summary>
        private bool _isDisposed = false;
        private readonly Timer statusCheckTimer;
        private GraphServiceClient _graphApiClient = null;

        private List<IParticipant> _noKickRetryUserList = new();

        // Key is: call ID + partipant ID
        private Dictionary<string, DateTime> _removeWarningsGivenCache = new();

        /// <summary>
        /// Initializes a new instance of the <see cref="CallHandler" /> class.
        /// </summary>
        /// <param name="statefulCall">The stateful call.</param>
        /// <param name="settings">The settings.</param>
        /// <param name="eventPublisher">The event publisher.</param>
        public CallHandler(
            ICall statefulCall,
            IAzureSettings settings,
            IEventPublisher eventPublisher
        )
            : base(TimeSpan.FromMinutes(10), statefulCall?.GraphLogger)
        {
            _settings = (AzureSettings)settings;
            _eventPublisher = eventPublisher;

            this.Call = statefulCall;
            this.Call.OnUpdated += this.CallOnUpdated;

            this.BotMediaStream = new BotMediaStream(this.Call.GetLocalMediaSession(), this.Call.Id, this.GraphLogger, eventPublisher, _settings);

            if (_settings.CaptureEvents)
            {
                var path = Path.Combine(Path.GetTempPath(), BotConstants.DefaultOutputFolder, _settings.EventsFolder, statefulCall.GetLocalMediaSession().MediaSessionId.ToString(), "participants");
                _capture = new CaptureEvents(path);
            }

            var confidentialClientApplication = ConfidentialClientApplicationBuilder
                .Create(_settings.AadAppId)
                .WithTenantId(_settings.AadTenantId)
                .WithClientSecret(_settings.AadAppSecret)
                .Build();

            _graphApiClient = new GraphServiceClient(new ClientCredentialProvider(confidentialClientApplication));



            // Initialize timer to check statuses
            var timer = new Timer(100 * 60); // every 60 seconds
            timer.AutoReset = true;
            timer.Elapsed += this.WebcamStatusCheck;
            this.statusCheckTimer = timer;

            Console.WriteLine($"Joining call ID {statefulCall.Id} on chat thread {statefulCall.Resource.ChatInfo.ThreadId}");
        }

        private void WebcamStatusCheck(object sender, ElapsedEventArgs e)
        {
            _ = Task.Run(async () =>
            {
                statusCheckTimer.Enabled = false;
                foreach (var p in this.Call.Participants)
                {
                    // Don't check your own (bot) webcam status
                    var participantIsThisBot = p.Resource?.Info?.Identity?.Application?.Id == _settings.AadAppId;
                    if (!participantIsThisBot)
                    {
                        var userHasWebcamOn = false;
                        var userStreams = ((Participant)((IResource)p).Resource).MediaStreams;
                        foreach (var s in userStreams)
                        {
                            if (s.MediaType.HasValue && s.MediaType.Value == Modality.Video && (s.Direction == MediaDirection.SendOnly || s.Direction == MediaDirection.SendReceive))
                            {
                                userHasWebcamOn = true;
                            }
                        }

                        // Find users without webcam on & that we haven't tried (and failed) to remove before
                        if (!userHasWebcamOn && !_noKickRetryUserList.Contains(p))
                        {
                            var userDisplayName = p.Resource?.Info?.Identity?.User?.DisplayName;
                            Console.WriteLine($"{userDisplayName} does not have webcam on");

                            // Have we warned this user for this call yet?
                            DateTime? lastBootWaring = UserWarned(this.Call.Id, p);

                            bool kickUser = lastBootWaring.HasValue && lastBootWaring.Value > DateTime.Now.AddMinutes(-5);
                            if (!kickUser)
                            {
                                // Warn to turn on webcam
                                var chatId = this.Call.Resource.ChatInfo.ThreadId;

                                // Doesn't work for bots joined by policy
                                await WarnUser(chatId, p);

                                // Next time they get kicked out the channel
                                SetUserHasBeenWarned(this.Call.Id, p.Id);
                            }
                            else
                            {
                                // User warned already; remove them from the call
                                try
                                {
                                    await p.DeleteAsync().ConfigureAwait(false);
                                }
                                catch (ServiceException ex)
                                {
                                    Console.WriteLine($"Couldn't remove {userDisplayName} - {ex.Message}");
                                    GraphLogger.Error(ex.Message);

                                    // Don't try to remove again
                                    _noKickRetryUserList.Add(p);
                                }
                            }

                        }
                    } // !participantIsThisBot
                }
                statusCheckTimer.Enabled = true;
            }).ForgetAndLogExceptionAsync(this.GraphLogger);
        }

        private async Task WarnUser(string chatId, IParticipant p)
        {

            var userName = p.Resource?.Info?.Identity?.User?.DisplayName;
            if (!string.IsNullOrEmpty(userName))
            {
                if (Call.Resource.State == CallState.Established)
                {
                    //await this.Call.PlayPromptAsync(new List<MediaPrompt> { warningMedia }).ConfigureAwait(false);
                    await Task.CompletedTask;
                }
            }
        }

        async Task<string> PostToCall(string url, string json)
        {
            var req = new HttpRequestMessage(HttpMethod.Post, url);
            await _graphApiClient.AuthenticationProvider.AuthenticateRequestAsync(req);

            using (var client = new HttpClient())
            using (var stringContent = new StringContent(json, Encoding.UTF8, "application/json"))
            {
                req.Content = stringContent;

                using (var response = await client
                    .SendAsync(req, HttpCompletionOption.ResponseHeadersRead)
                    .ConfigureAwait(false))
                {
                    var responseBody = await response.Content.ReadAsStringAsync();
                    response.EnsureSuccessStatusCode();
                    return responseBody;
                }
            }
        }


        private DateTime? UserWarned(string callId, IParticipant participant)
        {
            var key = callId + participant.Id;
            if (_removeWarningsGivenCache.ContainsKey(key))
            {
                return _removeWarningsGivenCache[key];
            }
            return null;
        }

        private void SetUserHasBeenWarned(string callId, string participantId)
        {
            var key = callId + participantId;
            if (_removeWarningsGivenCache.ContainsKey(key))
            {
                _removeWarningsGivenCache[key] = DateTime.Now;
            }
            else
            {
                _removeWarningsGivenCache.Add(key, DateTime.Now);
            }
        }


        /// <inheritdoc/>
        protected override Task HeartbeatAsync(ElapsedEventArgs args)
        {
            return this.Call.KeepAliveAsync();
        }

        /// <inheritdoc />
        protected override void Dispose(bool disposing)
        {

            base.Dispose(disposing);
            _isDisposed = true;
            this.Call.OnUpdated -= this.CallOnUpdated;

            this.BotMediaStream?.Dispose();

            this.statusCheckTimer.Enabled = false;

            // Event - Dispose of the call completed ok
            _eventPublisher.Publish("CallDisposedOK", $"Call.Id: {this.Call.Id}");
        }

        private void SetRecordingStatus(ICall source, ElapsedEventArgs e)
        {
            _ = Task.Run(async () =>
            {
                var newStatus = RecordingStatus.Recording;
                try
                {
                    // Event - Log the recording status
                    var status = Enum.GetName(typeof(RecordingStatus), newStatus);
                    _eventPublisher.Publish("SetRecordingStatus", $"Call.Id: {Call.Id} status changed to {status}");

                    // NOTE: if your implementation supports stopping the recording during the call, you can call the same method above with RecordingStatus.NotRecording
                    await source
                        .UpdateRecordingStatusAsync(newStatus)
                        .ConfigureAwait(false);

                }
                catch (Exception ex)
                {
                    // e.g. bot joins via direct join - may not have the permissions
                    GraphLogger.Error(ex, $"Failed to flip the recording status to {newStatus}");
                    // Event - Recording status exception - failed to update 
                    _eventPublisher.Publish("CallRecordingFlip", $"Failed to flip the recording status to {newStatus}");
                }

            }).ForgetAndLogExceptionAsync(this.GraphLogger);
        }


        /// <summary>
        /// Event fired when the call has been updated.
        /// </summary>
        /// <param name="sender">The call.</param>
        /// <param name="e">The event args containing call changes.</param>
        private async void CallOnUpdated(ICall sender, ResourceEventArgs<Call> e)
        {
            GraphLogger.Info($"Call status updated to {e.NewResource.State} - {e.NewResource.ResultInfo?.Message}");
            // Event - Recording update e.g established/updated/start/ended
            Console.WriteLine($"Call{e.NewResource.State}", $"Call.ID {Call.Id} Sender.Id {sender.Id} status updated to {e.NewResource.State} - {e.NewResource.ResultInfo?.Message}");

            if (e.OldResource.State != e.NewResource.State && e.NewResource.State == CallState.Established)
            {
                if (!_isDisposed)
                {
                    // await ConfigureCallSettings();

                    // Call is established. We should start receiving Audio, we can inform clients that we have started recording.
                    SetRecordingStatus(sender, null);

                    // Start tracking
                    this.statusCheckTimer.Enabled = true;
                }
            }

            if ((e.OldResource.State == CallState.Established) && (e.NewResource.State == CallState.Terminated))
            {
                if (BotMediaStream != null)
                {
                    var aQoE = BotMediaStream.GetAudioQualityOfExperienceData();

                    if (aQoE != null)
                    {
                        if (_settings.CaptureEvents)
                            await _capture?.Append(aQoE);
                    }
                    await BotMediaStream.StopMedia();
                }

                if (_settings.CaptureEvents)
                    await _capture?.Finalise();
            }
        }

    }
}
