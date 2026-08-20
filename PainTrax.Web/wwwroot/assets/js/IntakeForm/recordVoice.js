// Store recorders and URLs globally to handle multiple fields if necessary

let mediaRecorder;
let audioChunks = [];


async function startVoice(icon, ctrl) {
    var input = $('#' + ctrl);
    var iconElem = $(icon);

    if (!('webkitSpeechRecognition' in window)) {
        alert("Voice recognition not supported in this browser");
        return;
    }

    // 👉 STOP logic
    if (isRecording) {
        recognition.stop();
        if (mediaRecorder && mediaRecorder.state !== 'inactive') {
            mediaRecorder.stop();
        }
        return;
    }

    try {
        // 1. Get Microphone Access
        const stream = await navigator.mediaDevices.getUserMedia({ audio: true });

        // 2. Initialize MediaRecorder
        audioChunks = [];
        mediaRecorder = new MediaRecorder(stream);

        mediaRecorder.ondataavailable = (event) => {
            if (event.data.size > 0) audioChunks.push(event.data);
        };

        mediaRecorder.onstop = () => {
            const audioBlob = new Blob(audioChunks, { type: 'audio/webm' });
            const audioUrl = URL.createObjectURL(audioBlob);

            // Create or update the download button
             createDownloadButton(iconElem, ctrl, audioUrl);

            let formData = new FormData();
            formData.append('audio_data', audioBlob);
            formData.append('controlName', ctrl); // e.g., "StreetAddress"

            saveAudioFile(formData, ctrl);

            // Release microphone
            stream.getTracks().forEach(track => track.stop());
        };

        // 3. Initialize Speech Recognition
        recognition = new webkitSpeechRecognition();
        recognition.continuous = true;
        recognition.interimResults = false;
        recognition.lang = "en-US";

        recognition.onresult = function (event) {
            var transcript = convertPunctuation(
                event.results[event.results.length - 1][0].transcript
            );

            if (transcript.toLowerCase().trim() === "clear all") {
                if (confirm("Clear all text?")) {
                    input.val('');
                }
            } else {
                var currentText = input.val() || "";
                input.val((currentText + " " + transcript).trim());
            }
            if (typeof updatePreview === "function") updatePreview();
        };

        recognition.onend = function () {
            isRecording = false;
            iconElem.css("color", "#007bff").attr("title", "Start Recording");
        };

        // 4. Start both
        recognition.start();
        mediaRecorder.start();

        isRecording = true;
        iconElem.css("color", "red").attr("title", "Stop Recording");

        // Hide download button while recording a new one (optional)
        $(`#btn-dl-${ctrl}`).hide();

    } catch (err) {
        console.error("Mic error:", err);
        alert("Microphone access is required to record audio.");
    }
}

/**
 * Creates or updates a download button next to the mic icon
 */
function createDownloadButton(iconElem, ctrl, url) {
    const btnId = `btn-dl-${ctrl}`;
    let dlBtn = $(`#${btnId}`);

    // If button doesn't exist, create it
    if (dlBtn.length === 0) {
        dlBtn = $(`
            <a id="${btnId}" 
               class="btn btn-sm ml-2" 
               style="display: inline-flex; align-items: center; text-decoration: none;" 
               title="Download Audio">
               <i class="fa fa-download"></i> 
               <span style="margin-left: 5px; font-size: 12px;"></span>
            </a>
        `);
        // Insert it right after the mic icon
        iconElem.after(dlBtn);
    }

    // Update the link and filename
    const timestamp = new Date().toLocaleTimeString().replace(/[:\s]/g, '-');
    dlBtn.attr("href", url);
    dlBtn.attr("download", `recording-${ctrl}-${timestamp}.webm`);
    dlBtn.show(); // Ensure it's visible
}

function saveAudioFile(formData, ctrl) {

    $.ajax({
        url: saveAudioURl,
        type: 'POST',
        data: formData,
        processData: false,
        contentType: false,
        success: function (res) {
            if (res.success) {
                //// 1. Show the download/play button
                //showAudioButton(iconElem, ctrl, res.url);

                // 2. Set the filename in a hidden field
                // If ctrl is "StreetAddress", hidden field name should be "StreetAddressAudio"
                setHiddenFieldValue(ctrl, res.fileName);
            }
        }
    });
}

function setHiddenFieldValue(ctrl, fileName) {
    // We want the hidden field to be named after the control + "Audio"
    const hiddenName = ctrl + "Audio";

    // Check if the input already exists in the patientForm
    let hiddenInput = $('#patientForm input[name="' + hiddenName + '"]');

    if (hiddenInput.length === 0) {
        // 1. Create the input if it's not there
        // 2. Append it specifically to the #patientForm
        $('<input>').attr({
            type: 'hidden',
            id: hiddenName,
            name: hiddenName,
            value: fileName
        }).appendTo('#patientForm');

        console.log("Hidden field created and added to form.");
    } else {
        // 3. If it already exists, just update the value with the new recording filename
        hiddenInput.val(fileName);
        console.log("Hidden field updated.");
    }
}