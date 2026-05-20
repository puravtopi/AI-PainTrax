
var docTemplate = null;
var formData = null;
// Path to your saved template file (relative to this HTML file).
// (The DOCX is stored alongside this page.)


var recognition;
var isRecording = false;

function calculateBMI() {
    let weight = parseFloat($("#weight").val());
    let heightCm = parseFloat($("#height").val());

    if (!weight || !heightCm) {
        $("#result").html("Please enter valid values");
        return;
    }

    // Convert cm to meter
    let heightM = heightCm / 100;

    // BMI Formula
    let bmi = weight / (heightM * heightM);

    bmi = bmi.toFixed(2);

    $("#bmi").val(bmi)
}

function safeParse(json) {
    if (typeof json !== "string") return json;

    try {
        json = json.trim();
        json = json.replace(/^\uFEFF/, '');
        json = json.replace(/&quot;/g, '"');
        json = json.replace(/&#x2B;/g, '+');
        debugger
        return JSON.parse(json);
    } catch (e) {
        console.error("Invalid JSON:", json);
        throw e;
    }
}
function startVoice(icon, ctrl) {

    var input = $('#' + ctrl);

    if (!('webkitSpeechRecognition' in window)) {
        alert("Voice recognition not supported in this browser");
        return;
    }

    // 👉 STOP if already recording
    if (isRecording) {
        recognition.stop();
        isRecording = false;
        $(icon).css("color", "#007bff");
        $(icon).attr("title", "Start Recording");
        return;
    }

    // 👉 START recording
    recognition = new webkitSpeechRecognition();
    recognition.continuous = true;   // 🔥 IMPORTANT
    recognition.interimResults = false;
    recognition.lang = "en-US";

    recognition.start();
    isRecording = true;

    $(icon).css("color", "red");
    $(icon).attr("title", "Stop Recording");

    recognition.onresult = function (event) {

        var transcript = convertPunctuation(
            event.results[event.results.length - 1][0].transcript
        );

        if (transcript.toLowerCase() === "clear all") {
            if (confirm("Clear all text?")) {
                input.val('');
            }
        } else {
            var currentText = input.val() || "";
            input.val((currentText + "" + transcript).trim());
        }

        updatePreview();
    };

    recognition.onend = function () {
        isRecording = false;
        $(icon).css("color", "#007bff");
        $(icon).attr("title", "Stop Recording");
    };
}


function convertPunctuation(text) {
    return text
        .replace(/\bfull stop\b/gi, ".")
        .replace(/\bstop\b/gi, ".")
        .replace(/\bdot\b/gi, ".")
        .replace(/\bcomma\b/gi, ",")
        .replace(/\bquestion mark\b/gi, "?")
        .replace(/\bexclamation mark\b/gi, "!")
        .replace(/\bcolon\b/gi, ":")
        .replace(/\bsemicolon\b/gi, ";")
        .replace(/\bnew line\b/gi, "\n")
        .replace(/\bnext line\b/gi, "\n");
}
let selectedSections = [];
let selectedPESections = [];

function fnShowPresentComplain(cntrl, divCC, divPE) {

    if (!divCC) return;

    if ($(cntrl).is(":checked")) {

        // Add in order
        if (!selectedSections.includes(divCC)) {
            selectedSections.push(divCC);
        }

        if (divPE && !selectedPESections.includes(divPE)) {
            selectedPESections.push(divPE);
        }

    } else {

        // Remove when unchecked
        selectedSections = selectedSections.filter(x => x !== divCC);
        selectedPESections = selectedPESections.filter(x => x !== divPE);

        // Clear values
        $("#" + divCC).find("input").prop("checked", false).val("");
        $("#" + divPE).find("input").prop("checked", false).val("");
    }

    renderSectionsInOrder();
}

function renderSectionsInOrder() {

    let ccContainer = $("#presentComplaintContainer");   // create this div
    let peContainer = $("#physicalExamContainer");       // create this div

    ccContainer.empty();
    peContainer.empty();

    // Render CC sections in order
    selectedSections.forEach(function (id) {
        let el = $("#" + id);
        el.show();                 // ensure visible
        ccContainer.append(el);    // append in order
    });

    //Render PE sections in order
    selectedPESections.forEach(function (id) {
        let el = $("#" + id);
        el.show();
        peContainer.append(el);
    });
}

// CLOSE (❌)
$("#closePreview").click(function () {

    $("#previewSection").hide();

    // Expand form to full width
    $("#formSection")
        .removeClass("col-md-6")
        .addClass("col-md-12");

    // Show open button
    $("#openPreview").show();
});

// OPEN AGAIN
$("#openPreview").click(function () {

    $("#previewSection").show();

    // Back to split layout
    $("#formSection")
        .removeClass("col-md-12")
        .addClass("col-md-6");

    $(this).hide();
});


function loadTemplate() {
    // Try loading a DOCX template first; fall back to plain text.
    loadDocxTemplate(templateUrlDocx)
        .catch(function () {
            return loadTextTemplate(templateUrl);
        })
        .then(function () {
            updatePreview();
        });
}

function loadDocxTemplate(url) {
    return fetch(url)
        .then(function (res) {
            if (!res.ok) throw new Error("Failed to load DOCX template");
            return res.arrayBuffer();
        })
        .then(function (arrayBuffer) {
            return mammoth.convertToHtml({ arrayBuffer: arrayBuffer });
        })
        .then(function (result) {
            // Convert to HTML so formatting (bold, line breaks, etc.) is preserved.
            docTemplate = result.value;
        });
}

function loadTextTemplate(url) {
    return fetch(url)
        .then(function (res) {
            if (!res.ok) throw new Error("Failed to load text template");
            return res.text();
        })
        .then(function (text) {
            docTemplate = text;
        });
}

function applyTemplate(template, values) {
    return template.replace(/{{\s*([^}]+)\s*}}/g, function (_, key) {
        return values[key] || "";
    });
}

function updatePreview() {

    var values = {};

    $('#patientForm').find('input, select, textarea').each(function () {
        var key = $(this).attr('name') || $(this).attr('id');
        var cls = $(this).attr('class');
        if (!key) return;

        var type = $(this).attr('type');

        if (type === 'radio') {
            if ($(this).is(':checked')) {
                values[key] = $(this).val();
            }
        }
        else if (type === 'checkbox') {

            // 🔹 Special handling for Complaints (multi-checkbox)
            if (cls === "chkMulti") {
                if (!values[key]) values[key] = [];
                if ($(this).is(':checked')) {
                    values[key].push($(this).val());
                }
            }
            else {
                // normal checkbox (true/false)
                values[key] = $(this).is(':checked');
            }
        }
        else {
            values[key] = $(this).val() || "";
        }
    });

    values["Id"] = $("#hdId").val();

    var today = new Date();

    var options = {
        month: 'short',
        day: '2-digit',
        year: 'numeric'
    };

    // var formattedDate = today.toLocaleDateString('en-US', options);

    // values["DOE"] = formattedDate.replace(',', '');

    if ($("#DOE").val() != '')
        values["DOE"] = formatDate($("#DOE").val());

    if ($("#DOB").val() != '')
        values["DOB"] = formatDate($("#DOB").val());

    if ($("#DOA").val() != '') {
        $("#lblDOA").val($("#DOA").val());
        values["DOA"] = formatDate($("#DOA").val());
    }

    formData = values;

    var template = docTemplate ||
        "Patient {{PatientName}} ({{Gender}}), dominant hand {{DominantHand}},\n" +
        "reports that on {{DOA}}, while working as {{JobTitle}},\n" +
        "they sustained an injury when {{Incident}}.\n\n" +
        "The incident occurred {{IncidentTypes}} and involved {{BodyPart}}.\n" +
        "Since the incident, symptoms have been persistent and\n" +
        "interfere with daily activities.";

    var output = applyTemplate(template, values);
    $("#previewText").html(output);

    processTemplate(values)
}

// Trigger on all inputs
$(document).on("keyup change", "#patientForm input, #patientForm select, .incident-type-option", function () {
    updatePreview();
});

function processTemplate(data) {
    let html = $('#previewText').html();

    // Replace placeholders
    Object.keys(data).forEach(key => {
        let regex = new RegExp('{{' + key + '}}', 'g');
        html = html.replace(regex, data[key]);
    });

    if (data.PreInjury === "No")
        $("#preInjuryDetailsBox").show();
    else
        $("#preInjuryDetailsBox").hide();

    if (data.Hospital === "Yes")
        $("#divHospital").show();
    else
        $("#divHospital").hide();

    if (data.Hospital === "Yes")
        $("#divHospital").show();
    else
        $("#divHospital").hide();

    if (data.Allergy === "Yes")
        $("#divAllergy").show();
    else
        $("#divAllergy").hide();

    if (data.Smoke === "Yes")
        $("#divSmoke").show();
    else
        $("#divSmoke").hide();

    if (data.OtherDrugs === "Yes")
        $("#divDrugs").show();
    else
        $("#divDrugs").hide();

    if (data.PT === "Yes")
        $("#divPT").show();
    else
        $("#divPT").hide();

    if (data.Chiro === "Yes")
        $("#divChiro").show();
    else
        $("#divChiro").hide();

    $(".mmClass").show();
    if (data.InjuryType === "WC") {
        $("#divWC").show();
        $("#divWC1").show();
        $("#divNF").hide();
        $("#divPI").hide();
        html = html.replace('[WC_START]', '').replace('[WC_END]', '');
        html = html.replace(/\[MVA_START\][\s\S]*?\[MVA_END\]/g, '');
    }
    else if (data.InjuryType === "NF") {
        // Keep MVA, remove WC
        html = html.replace('[MVA_START]', '').replace('[MVA_END]', '');
        html = html.replace(/\[WC_START\][\s\S]*?\[WC_END\]/g, '');
        $("#divWC").hide();
        $("#divWC1").hide();
        $("#divNF").show();
        $("#divPI").hide();
    }
    else if (data.InjuryType === "MM") {
        // $("#divNF").hide();
        // $("#divMM").hide();
        // $("#divWC1").hide();
        $(".mmClass").hide();
        $("#divPI").hide();
    }
    else if (data.InjuryType === "PI") {
        $("#divWC").hide();
        $("#divWC1").hide();
        $("#divNF").hide();
        $("#divPI").show();
    }
    else {
        // Remove both
        html = html.replace(/\[WC_START\][\s\S]*?\[WC_END\]/g, '');
        html = html.replace(/\[MVA_START\][\s\S]*?\[MVA_END\]/g, '');
        $("#divNF").hide();
        $("#divMM").hide();
        $("#divWC1").hide();
        $("#divPI").hide();
    }

    $('#previewText').html(html);
}

function saveForm(isSubmit) {
    formData["Diagnosis"] = $("#divDiagnosis").html();
    formData["TreatmentIds"] = $("#TreatmentIds").val();
    formData["TreatmentDesc"] = $("#TreatmentDesc").val();
    formData["TreatmentDelimitDesc"] = $("#TreatmentDelimitDesc").val();
    // Serialize form data
    var url = '@Url.Action("SaveForm", "IntakeForm")';
    var indexurl = '@Url.Action("Index", "Visit")';
    var data = JSON.stringify(formData);

    $.ajax({
        url: url, // Update with your Controller/Action
        type: 'POST',
        contentType: "application/json",
        data: data,
        success: function (response) {
            if (response.success) {
                alert(response.message);
                // // Optionally reset form
                $('#patientForm')[0].reset();
                // updatePreview();
                if (isSubmit)
                    window.location.href = indexurl;
                else
                    window.location.href = '@Url.Action("AIInitialIntake", "IntakeForm")?id=' + response.id + '&locid=' + response.locid;
            } else {
                alert("Error: " + response.message);
            }
        },

        error: function () {
            alert("An error occurred while saving data.");
        }
    });
}

$(document).ready(function () {
    // Load template and initialize preview
    var today = new Date().toISOString().split('T')[0];
    $('#DOE').val(today);
    loadTemplate();
    if (formData != '') {
        bindForm(formData);
    }

    $('#patientForm').on('submit', function (e) {

        if (!checkValidation())
            return false;

        e.preventDefault(); // Stop page reload
        saveForm(true);

    });



    var selected = [];
    $(document).on('change', 'input[name="Complaints"]', function () {

        var section = $(this).data('section');
        var pe = $(this).data('pe');

        fnShowPresentComplain(this, section, pe);



        // $('input[name="Complaints"]:checked').each(function () {
        //     selected.push($(this).val());
        // });
        var value = $(this).val();



        if ($(this).is(':checked')) {
            // Add in click order
            selected.push(value);
        } else {
            // Remove if unchecked
            selected = selected.filter(function (item) {
                return item !== value;
            });
        }

        if (selected.length > 0) {

            if (selected.includes("Neck")) {
                $(".upper").show();
            }
            else {
                $(".upper").hide();
            }

            if (selected.includes("Lowback")) {
                $(".lower").show();
            }
            else {
                $(".lower").hide();
            }



            $("#acordPresentComp").show();
            $("#acordPhyExam").show();
            $("#acordDiagnosis").show();


            var html = '';

            selected.forEach(function (part) {
                part = part.trim();

                html += `<span class="diagnosis-link"
                                                data-part="${part}"
                                                style="margin-left:8px; color:blue; cursor:pointer; text-decoration:underline;">
                                            ${part}
                                         </span>`;
            });

            $('#bodyPartLinks').html(html);
        }
        else {
            $("#acordPresentComp").hide();
            $("#acordPhyExam").hide();
            $("#acordDiagnosis").hide();
        }

        var result = selected.join(', ');

        console.log(result);

        $('#BodyPart').val(result);
    });

    $('#DOB').on('change', function () {
        var dob = $(this).val(); // format: yyyy-mm-dd
        var age = calculateAge(dob);
        $('#Age').val(age);
    });
});

function fnExportDOC() {
    var htmlContent = $('#previewText').html();
    var url = '@Url.Action("ExportWord", "IntakeForm")';
    var form = $('<form method="post" action="' + url + '"></form>');

    form.append('<input type="hidden" name="htmlContent" value="'
        + $('<div>').text(htmlContent).html() + '">');

    $('body').append(form);
    form.submit();
}

function bindForm(val) {
    // var cleaned = val
    //     .replace(/&quot;/g, '"')   // decode HTML quotes
    //     .trim();

    // cleaned=cleaned.replace(/"/g, '\\"');

    //
    var data = safeParse(val);


    $('#patientForm').find('input, select, textarea').each(function () {

        var key = $(this).attr('name') || $(this).attr('id');
        if (!key || data[key] === undefined || data[key] === null) return;

        var type = $(this).attr('type');
        var value = data[key];

        // 🔹 Fix DATE format
        if (type === 'date') {
            value = formatToISODate(value);
        }

        if (type === 'radio') {
            if ($(this).val() == value) {
                $(this).prop("checked", true);
            }
        }
        else if (type === 'checkbox') {

            // 🔹 Handle multi-checkbox (like Complaints array)
            if (Array.isArray(value)) {
                if (value.includes($(this).val())) {
                    $(this).prop("checked", true);
                } else {
                    $(this).prop("checked", false);
                }
            }
            else {
                // normal checkbox (true/false)
                $(this).prop("checked", value === true || value === "true");
            }
        }
        else {
            $(this).val(value);
        }

    });


    $("#hdId").val(Id);


    setTimeout(function () {
        $('#patientForm').find('input[name="Complaints"]:checked').trigger('change');
        updatePreview();
    }, 1000);
}

function formatDate(inputDate) {

    var date = inputDate.split('-');

    // Extract date components
    var month = parseInt(date[1]); // Months are zero based
    var day = parseInt(date[2])
    var year = date[0];

    // Pad single digit month and day with leading zeros
    if (month < 10) {
        month = "0" + month;
    }
    if (day < 10) {
        day = "0" + day;
    }

    // Return formatted date
    return month + "/" + day + "/" + year;
}

function formatToISODate(dateStr) {
    if (!dateStr) return "";

    // Already correct
    if (/^\d{4}-\d{2}-\d{2}$/.test(dateStr)) return dateStr;

    let date = new Date(dateStr);

    if (isNaN(date)) return "";

    let year = date.getFullYear();
    let month = String(date.getMonth() + 1).padStart(2, '0');
    let day = String(date.getDate()).padStart(2, '0');

    return `${year}-${month}-${day}`;
}

function calculateAge(dob) {
    var birthDate = new Date(dob);
    var today = new Date();

    var age = today.getFullYear() - birthDate.getFullYear();
    var monthDiff = today.getMonth() - birthDate.getMonth();

    // Adjust if birthday hasn't occurred yet this year
    if (monthDiff < 0 || (monthDiff === 0 && today.getDate() < birthDate.getDate())) {
        age--;
    }

    return age;
}

function fnshowDLPopup() {
    $("#fuModal").modal('show');
}

function fnshowTreatment() {
    $("#treatmentModal").modal('show');
}

function fnTreatmentDetails() {
    var strContent = '<ol>';
    var strIds = '', strDesc = '';

    $('input[type=checkbox].treatment').each(function (i, e) {
        // Check if the checkbox is selected
        if (e.checked) {
            var checkboxValue = $("#chk_" + e.id).val();

            // Only append to strContent if checkboxValue is not undefined or empty
            if (checkboxValue) {
                strContent += '<li>' + checkboxValue + '</li>';
                strIds += (strIds ? ',' : '') + e.id;  // Ensure no leading comma
                strDesc += (strDesc ? '^' : '') + checkboxValue;  // Ensure no leading ^ for first value
            }
        }
    });

    // Close the ordered list
    strContent += '</ol>';

    // If no valid checkboxes were selected, clear the content (optional)
    if (strContent === '<ol></ol>') {
        strContent = '';  // Or handle this case as needed
    }

    //var cmp_id = '@ViewBag.CmpId';

    // if (cmp_id === '18' || cmp_id === '15' || cmp_id === '10' || cmp_id === '2') {

    //     strContent = strContent.replace(/<ol>/g, '')     // remove opening <ol>
    //         .replace(/<\/ol>/g, '')   // remove closing </ol>
    //         .replace(/<li>/g, '')     // remove opening <li>
    //         .replace(/<\/li>/g, ''); // replace closing </li> with <br/>;
    // }

    $("#divTreatment").html(strContent);
    $('#TreatmentIds').val(strIds);
    $('#TreatmentDelimitDesc').val(strDesc);
    $('#TreatmentDesc').val(strContent);
}

$(document).on('click', '.diagnosis-link', function (e) {
    e.stopPropagation(); // 🔥 prevents accordion from opening/closing

    var part = $(this).data('part');
    fnShowDiagnosis(part);
});

function fnShowDiagnosis(bodyParts) {

    var url = '@Url.Action("GetDaignoCodeList", "IntakeForm")?bodyparts=' + bodyParts + "&id=" + $("#hdId").val();
    $("#hdSelectedBody").val(bodyParts);
    $.ajax({
        type: "Post",
        url: url,

        contentType: "application/x-www-form-urlencoded",
        success: function (data, status, xhr) {
            debugger
            $('#modelDaignoCode').html(data);
        },
        error: function (xhr, status, error) {
            alert("Error!" + error);
        },
    });

    $('#daignoCodeModal').modal('show');
    return false;
}

function fnUploadDL() {

    //var fileInput = $('#fup')[0];

    if (filesArr.length === 0) {
        alert("Please select an image");
        return;
    }

    var formData = new FormData();
    formData.append("file", filesArr[0]);

    var url = '@Url.Action("UploadDL", "IntakeForm")';

    // ✅ Show loader inside modal
    $('#uploadLoader').addClass('show');

    $.ajax({
        url: url,
        type: 'POST',
        data: formData,
        contentType: false,
        processData: false,

        success: function (response) {

            let data = response.parsedData;

            $('#FN').val(data.firstName);
            $('#LN').val(data.lastName);
            $('#Gender').val(data.gender);
            $('#DLPath').val(data.fileName);

            setDateToInput(data.dob, "DOB");
            $('#uploadLoader').addClass('hide');
            $("#fuModal").modal('hide');
            $("#previewGrid").html('');

        },

        error: function (err) {
            console.log(err);
            alert("Something went wrong!");
        },

        complete: function () {
            // ✅ Hide loader
            $('#uploadLoader').removeClass('show');
        }
    });
}

function setDateToInput(dateStr, inputId) {

    if (!dateStr) return;

    // Normalize separator
    dateStr = dateStr.replace(/-/g, "/");

    var parts = dateStr.split("/");

    if (parts.length !== 3) return;

    var mm = parts[0].padStart(2, '0');
    var dd = parts[1].padStart(2, '0');
    var yy = parts[2];

    // 🔥 Handle 2-digit year
    if (yy.length === 2) {
        var currentYear = new Date().getFullYear() % 100;

        // Rule: 00–currentYear → 2000s, else 1900s
        yy = (parseInt(yy) <= currentYear ? '20' : '19') + yy;
    }

    var formattedDate = `${yy}-${mm}-${dd}`;

    // Set DOB
    $('#' + inputId).val(formattedDate);

    // Set Age
    $('#Age').val(calculateAge(formattedDate));
}


let filesArr = [];

const dropArea = document.getElementById("dropArea");
const fileInput = document.getElementById("fup");
const previewGrid = document.getElementById("previewGrid");

// Click
dropArea.onclick = () => fileInput.click();

// Select
fileInput.onchange = (e) => handleFiles(e.target.files);

// Drag
dropArea.ondragover = e => e.preventDefault();
dropArea.ondrop = e => {
    e.preventDefault();
    handleFiles(e.dataTransfer.files);
};

// Handle files
function handleFiles(files) {
    [...files].forEach(file => {
        file.id = Date.now() + Math.random();
        filesArr.push(file);
        renderPreview(file);
    });
}

// Preview
function renderPreview(file) {
    const reader = new FileReader();

    reader.onload = e => {
        const div = document.createElement("div");
        div.className = "preview-item";
        div.dataset.id = file.id;

        div.innerHTML = `
    <button class="remove-btn">×</button>
    <img src="${e.target.result}" />
    <div class="progress"><div class="progress-bar"></div></div>
    `;

        // Remove
        div.querySelector(".remove-btn").onclick = () => {
            filesArr = filesArr.filter(f => f.id !== file.id);
            div.remove();
        };

        previewGrid.appendChild(div);
    };

    reader.readAsDataURL(file);
}

// Reorder
new Sortable(previewGrid, {
    animation: 150,
    onEnd: () => {
        let newOrder = [];
        document.querySelectorAll(".preview-item").forEach(item => {
            let id = item.dataset.id;
            let file = filesArr.find(f => f.id == id);
            if (file) newOrder.push(file);
        });
        filesArr = newOrder;
    }
});

// Upload
document.getElementById("uploadBtn").onclick = function (e) {

    e.preventDefault();

    if (!filesArr || filesArr.length === 0) {
        alert("Please select files first.");
        return;
    }

    const progressBars = document.querySelectorAll(".progress-bar");

    filesArr.forEach((file, index) => {

        let formData = new FormData();
        formData.append("file", file);
        formData.append("order", index);

        let xhr = new XMLHttpRequest();

        xhr.open("POST", "/Media/UploadNew", true);

        // Upload progress
        xhr.upload.onprogress = function (e) {

            if (e.lengthComputable && progressBars[index]) {

                let percent = Math.round((e.loaded / e.total) * 100);

                progressBars[index].style.width = percent + "%";
                progressBars[index].innerText = percent + "%";
            }
        };

        // Success
        xhr.onload = function () {

            if (xhr.status === 200) {
                console.log("Uploaded:", file.name);
            }
            else {
                console.error("Upload failed:", file.name);
            }
        };

        // Error
        xhr.onerror = function () {
            console.error("Network error while uploading:", file.name);
        };

        xhr.send(formData);
    });
};


function hasSectionValue(sectionId) {

    let hasValue = false;

    $("#" + sectionId)
        .find("input, textarea, select")
        .each(function () {

            const type = $(this).attr("type");

            // checkbox / radio
            if ((type === "checkbox" || type === "radio") && $(this).is(":checked")) {
                hasValue = true;
                return false;
            }

            // text / textarea / select
            if (type !== "checkbox" && type !== "radio") {

                const value = $(this).val();

                if (value != null && value.toString().trim() !== "") {
                    hasValue = true;
                    return false;
                }
            }
        });

    return hasValue;
}

function validateBodySection(bodyPart, complaintSectionId, peSectionId) {

    const isSelected = $(`input[name='Complaints'][value='${bodyPart}']`).is(":checked");

    if (!isSelected)
        return true;

    if (!hasSectionValue(complaintSectionId)) {
        alert(`Please enter ${bodyPart} complaint before submitting.`);
        return false;
    }

    if (!hasSectionValue(peSectionId)) {
        alert(`Please enter ${bodyPart} physical exam details before submitting.`);
        return false;
    }

    return true;
}

function checkValidation() {

    const validations = [
        {
            bodyPart: "Neck",
            complaintSectionId: "neckSection",
            peSectionId: "neckPE"
        },
        {
            bodyPart: "Midback",
            complaintSectionId: "mbSection",
            peSectionId: "mbPE"
        },
        {
            bodyPart: "Lowback",
            complaintSectionId: "lbSection",
            peSectionId: "lbPE"
        },
        {
            bodyPart: "Right Shoulder",
            complaintSectionId: "rshSection",
            peSectionId: "rshPE"
        },
        {
            bodyPart: "Left Shoulder",
            complaintSectionId: "lshSection",
            peSectionId: "lshPE"
        },
        {
            bodyPart: "Right Knee",
            complaintSectionId: "rknSection",
            peSectionId: "rknPE"
        },
        {
            bodyPart: "Left Knee",
            complaintSectionId: "lknSection",
            peSectionId: "lknPE"
        }
    ];

    for (const item of validations) {

        const isValid = validateBodySection(
            item.bodyPart,
            item.complaintSectionId,
            item.peSectionId
        );

        if (!isValid)
            return false;
    }

    return true;
}

async function downloadPDF() {

    $('#pdfLoader').show();
    try {
        // Store current accordion state
        let accordionStates = [];

        $('.accordion-collapse').each(function () {
            accordionStates.push($(this).hasClass('show'));
        });

        // Open all accordions
        $('.accordion-collapse').addClass('show');

        // Hide accordion buttons/icons if needed
        $('.accordion-button').hide();

        // Hide microphone buttons
        $('.mic-btn').hide();
        $('#btnSection').hide();

        // Optional: remove borders/shadows for clean PDF
        $('.accordion-item').css({
            'border': 'none',
            'box-shadow': 'none'
        });

        //         $('#patientForm').css({
        //     'font-size': '11px',
        //     'line-height': '1.3'
        // });

        // $('#patientForm input, #patientForm textarea, #patientForm select, #patientForm p, #patientForm label').css({
        //     'font-size': '11px',
        //     'padding': '2px 4px',
        //     'height': 'auto'
        // });

        // $('h4, h5, h6').css({
        //     'font-size': '14px'
        // });

        // Wait for rendering
        await new Promise(resolve => setTimeout(resolve, 500));

        // Target form/div
        const element = document.getElementById('patientForm');

        // PDF options
        const opt = {
            margin: 0.3,
            filename: 'Patient_Injury_Report.pdf',
            image: {
                type: 'jpeg',
                quality: 1
            },
            html2canvas: {
                scale: 2,
                useCORS: true,
                scrollY: 0
            },
            jsPDF: {
                unit: 'in',
                format: 'a4',
                orientation: 'portrait'
            },
            pagebreak: {
                mode: ['avoid-all', 'css', 'legacy']
            }
        };

        // Generate PDF
        await html2pdf().set(opt).from(element).save();

        // Restore accordion state
        $('.accordion-collapse').each(function (index) {

            if (!accordionStates[index]) {
                $(this).removeClass('show');
            }
        });

        //         $('#patientForm').css({
        //     'font-size': '',
        //     'line-height': ''
        // });

        // $('#patientForm input, #patientForm textarea, #patientForm select,#patientForm p, #patientForm label').css({
        //     'font-size': '',
        //     'padding': '',
        //     'height': ''
        // });

        // $('h4, h5, h6').css({
        //     'font-size': ''
        // });

        // Show hidden buttons again
        $('.accordion-button').show();
        $('.mic-btn').show();
        $('#btnSection').show();
    }
    finally {

        // HIDE LOADER
        $('#pdfLoader').hide();
    }
}

