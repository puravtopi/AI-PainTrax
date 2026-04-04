using System.ComponentModel.DataAnnotations;
namespace PainTrax.Web.Models
{
    public class IntakeForm
    {
        // 1.Initial Evaluation Intake sheet
        public string? visit_type { get; set; }
        public string? location_id { get; set; }
        public string? l_name { get; set; }
        public string? f_name { get; set; }
        public string CaseType { get; set; }
        //public string? case_type_WC { get; set; }
        //public string? case_type_Lien { get; set; }
        public DateTime? dob { get; set; }
        public DateTime? doa { get; set; }
        public DateTime? dov { get; set; }
        public string? wcb { get; set; }
        public int? age { get; set; }
        public string? gender { get; set; }
        //public string? female { get; set; }
        public string? handedness { get; set; }
        //public string? handedness_L { get; set; }
        public string? occupation { get; set; }
        public string? still_working { get; set; }
        //public string? still_working_no { get; set; }
        public string? txt_still_working { get; set; }
        public string? degree_of_disability { get; set; }
        public DateTime? last_day_worked { get; set; }
        public DateTime? return_to_work { get; set; }
        public int? weight { get; set; }
        public int? height { get; set; }
        public int? temperature { get; set; }

        // 2.Description of the accident
        public string? description_of_the_accident { get; set; }

        //3.History given by patient
        public string? direct_impact_on_bodypart { get; set; }
        //public string? direct_impact_on_bodypart_no { get; set; }
        public string? if_yes_which_bodypart { get; set; }
        public string? loss_consciousness { get; set; }
        //public string? loss_consciousness_no { get; set; }
        public string? if_yes_how_long { get; set; }
        public string? did_go_to_the_hospital { get; set; }
        //public string? did_go_to_the_hospital_no { get; set; }
        public string? if_yes_which_hospital { get; set; }
        public string? go_to_the_sameday { get; set; }
        //public string? go_to_the_sameday_no { get; set; }
        public string? if_no_how_many_days_later { get; set; }
        public string? rdo_ambulance { get; set; }
        //public string? rdo_taxi { get; set; }
        //public string? rdo_family_friend { get; set; }
        //public string? rdo_alone { get; set; }
        public string? did_they_did_any_test { get; set; }
        //public string? did_they_did_any_test_no { get; set; }

        public string? what_test { get; set; }
        public string? what_tests_xray { get; set; }
        public string? what_tests_ct { get; set; }
        public string? what_tests_mri { get; set; }
        public string? body_parts { get; set; }
        public string? prescribed_any_medication_yes { get; set; }
        //public string? prescribed_any_medication_no { get; set; }
        public string? what_medication { get; set; }
        public string? involved_accident_past_yes { get; set; }
        //public string? involved_accident_past_no { get; set; }
        public string? when_accident_past { get; set; }
        public string? bodyparts_injured_inaccident { get; set; }
        public string? completly_recovered_yes { get; set; }
        //public string? completly_recovered_no { get; set; }
        public string? any_medical_conditions { get; set; }
        public string? any_medical_conditions_Diabeties { get; set; }
        public string? any_medical_conditions_bp { get; set; }
        public string? any_medical_conditions_ashthma { get; set; }
        public string? any_medical_conditions_heart { get; set; }
        public string? any_medical_conditions_none { get; set; }
        public string? medical_conditions_others { get; set; }
        public string? had_any_surgeries { get; set; }
        //public string? had_any_surgeries_no { get; set; }
        public string? had_any_surgeries_when { get; set; }
        public string? type_of_surgery { get; set; }
        public string? taking_medications_surgery { get; set; }
        // public string? taking_medications_surgery_no { get; set; }
        public string? what_medications_surgery { get; set; }
        public string? allergies { get; set; }
        // public string? allergies_no { get; set; }
        public string? what_allergies { get; set; }
        public string? family_history { get; set; }
        public string? smoke { get; set; }
        //public string? smoke_no { get; set; }
        public string? txt_smoke { get; set; }
        public string? marijuana { get; set; }
        //public string? marijuana_no { get; set; }
        public string? txt_marijuana { get; set; }
        public string? alcohol { get; set; }
        public string? txt_alcohol { get; set; }
        public string? social_history { get; set; }
        public string? physical_therapy { get; set; }
        //public string? physical_therapy_no { get; set; }
        public string? txt_physical_therapy { get; set; }
        public string? chiropractic { get; set; }
        //public string? chiropractic_no { get; set; }
        public string? txt_chiropractic { get; set; }
        public string? symptoms_of_accident_Headaches { get; set; }
        public string? symptoms_of_accident_ChestPain { get; set; }
        public string? symptoms_of_accident_Abdominal { get; set; }
        public string? symptoms_of_accident_Muscle { get; set; }
        public string? symptoms_of_accident_Dizziness { get; set; }
        public string? symptoms_of_accident_Nausea { get; set; }
        public string? symptoms_of_accident_Ringing_in_ears { get; set; }
        public string? symptoms_of_accident_Bladder { get; set; }
        public string? symptoms_of_accident_Bowel { get; set; }
        public string? symptoms_of_accident_Seizure { get; set; }
        public string? symptoms_of_accident_Sleep_issues { get; set; }
        public string? symptoms_of_accident_Anxiety { get; set; }

        public string? symptoms_since_accident { get; set; }
        //4.ADL ( Activities of daily living )
        //public string? personal_hygiene { get; set; }
        //public string? eating_independently { get; set; }
        //public string? Shop_for_food { get; set; }
        //public string? walking { get; set; }
        //public string? cooking { get; set; }
        //public string? sleeping { get; set; }
        public string? adl_personal { get; set; }
        public string? adl_eating { get; set; }
        public string? adl_Shop_for_food { get; set; }
        public string? adl_walking { get; set; }
        public string? adl_cooking { get; set; }
        public string? adl_sleeping { get; set; }

        // 5.How would you describe your pain?

        public string? txt_describe_neck { get; set; }
        public string? describe_neck_Constant { get; set; }
        public string? describe_neck_Intermittent { get; set; }
        public string? describe_neck_Sharp { get; set; }
        public string? describe_neck_Electric { get; set; }
        public string? describe_neck_Shooting { get; set; }
        public string? describe_neck_Throbbing { get; set; }
        public string? describe_neck_Pulsating { get; set; }
        public string? describe_neck_Dull { get; set; }
        public string? describe_neck_Achy { get; set; }
        public string? neck_pain_radiates_RUE { get; set; }
        public string? neck_pain_radiates_LUE { get; set; }
        public string? neck_pain_radiates_BUE { get; set; }
        public string? cc_neck { get; set; }

        public string? neck_pain_bodypart_shoulder { get; set; }
        public string? neck_pain_bodypart_elbow { get; set; }
        //public string? neck_pain_bodypart_arm { get; set; }
        //public string? neck_pain_bodypart_forearm { get; set; }
        public string? neck_pain_bodypart_hand { get; set; }
        public string? neck_pain_bodypart_wrist { get; set; }
        public string? neck_pain_bodypart_finger { get; set; }
        public string? neck_pain_numbness { get; set; }
        public string? neck_pain_tingling { get; set; }
        public string? neck_pain_bodypart_arm_1 { get; set; }
        public string? neck_pain_bodypart_forearm_1 { get; set; }
        public string? neck_pain_bodypart_hand_1 { get; set; }
        public string? neck_pain_bodypart_wrist_1 { get; set; }
        public string? neck_pain_bodypart_finger_1 { get; set; }
        public string? txt_describe_midback { get; set; }
        public string? describe_midback_Constant { get; set; }
        public string? describe_midback_Intermittent { get; set; }
        public string? describe_midback_Sharp { get; set; }
        public string? describe_midback_Electric { get; set; }
        public string? describe_midback_Shooting { get; set; }
        public string? describe_midback_Throbbing { get; set; }

        public string? cc_midback { get; set; }

        public string? describe_midback_Pulsating { get; set; }
        public string? describe_midback_Dull { get; set; }
        public string? describe_midback_Achy { get; set; }
        public string? txt_describe_lowback { get; set; }
        public string? describe_lowback_Constant { get; set; }
        public string? describe_lowback_Intermittent { get; set; }
        public string? describe_lowback_Sharp { get; set; }
        public string? describe_lowback_Electric { get; set; }
        public string? describe_lowback_Shooting { get; set; }
        public string? describe_lowback_Throbbing { get; set; }
        public string? describe_lowback_Pulsating { get; set; }
        public string? describe_lowback_Dull { get; set; }
        public string? describe_lowback_Achy { get; set; }

        public string? cc_lowback { get; set; }
        public string? lowback_pain_radiates_RLE { get; set; }
        public string? lowback_pain_radiates_LLE { get; set; }
        public string? lowback_pain_radiates_BLE { get; set; }
       
        public string? lowback_pain_bodypart_hip { get; set; }
        public string? lowback_pain_bodypart_thigh { get; set; }
        public string? lowback_pain_bodypart_knee { get; set; }
        public string? lowback_pain_bodypart_leg { get; set; }
        public string? lowback_pain_bodypart_ankle { get; set; }
        public string? lowback_pain_bodypart_foot { get; set; }
        public string? lowback_pain_bodypart_toe { get; set; }

        public string? lowback_pain_numbness { get; set; }
        public string? lowback_pain_tingling { get; set; }
        
        public string? lowback_pain_bodypart_numbness_thigh { get; set; }
        public string? lowback_pain_bodypart_numbness_knee { get; set; }
        public string? lowback_pain_bodypart_numbness_leg { get; set; }
        public string? lowback_pain_bodypart_numbness_ankle { get; set; }
        public string? lowback_pain_bodypart_numbness_foot { get; set; }

        public string? lowback_pain_bodypart_numbness_toe { get; set; }
        public string? describe_leftshoulder { get; set; }
        public string? txt_describe_leftshoulder_Constant { get; set; }
        public string? txt_describe_leftshoulder_Intermittent { get; set; }
        public string? txt_describe_leftshoulder_Sharp { get; set; }
        public string? txt_describe_leftshoulder_Electric { get; set; }
        public string? txt_describe_leftshoulder_Shooting { get; set; }
        public string? txt_describe_leftshoulder_Throbbing { get; set; }
        public string? txt_describe_leftshoulder_Pulsating { get; set; }
        public string? txt_describe_leftshoulder_Dull { get; set; }
        public string? txt_describe_leftshoulder_Achy { get; set; }

        public string? cc_l_shoulder { get; set; }

        public string? describe_rightshoulder { get; set; }
        public string? txt_describe_rightshoulder_Constant { get; set; }
        public string? txt_describe_rightshoulder_Intermittent { get; set; }
        public string? txt_describe_rightshoulder_Sharp { get; set; }
        public string? txt_describe_rightshoulder_Electric { get; set; }
        public string? txt_describe_rightshoulder_Shooting { get; set; }
        public string? txt_describe_rightshoulder_Throbbing { get; set; }
        public string? txt_describe_rightshoulder_Pulsating { get; set; }
        public string? txt_describe_rightshoulder_Dull { get; set; }

        public string? txt_describe_rightshoulder_Achy { get; set; }


        public string? cc_r_shoulder { get; set; }

        public string? describe_leftknee { get; set; }
        public string? txt_describe_leftknee_Constant { get; set; }
        public string? txt_describe_leftknee_Intermittent { get; set; }
        public string? txt_describe_leftknee_Sharp { get; set; }
        public string? txt_describe_leftknee_Electric { get; set; }
        public string? txt_describe_leftknee_Shooting { get; set; }
        public string? txt_describe_leftknee_Throbbing { get; set; }
        public string? txt_describe_leftknee_Pulsating { get; set; }

        public string? cc_l_knee { get; set; }

        public string? txt_describe_leftknee_Dull { get; set; }
        public string? txt_describe_leftknee_Achy { get; set; }
        public string? describe_rightknee { get; set; }
        public string? txt_describe_rightknee_Constant { get; set; }
        public string? txt_describe_rightknee_Intermittent { get; set; }
        public string? txt_describe_rightknee_Sharp { get; set; }
        public string? txt_describe_rightknee_Electric { get; set; }
        public string? txt_describe_rightknee_Shooting { get; set; }
        public string? txt_describe_rightknee_Throbbing { get; set; }
        public string? txt_describe_rightknee_Pulsating { get; set; }

        public string? txt_describe_rightknee_Dull { get; set; }
        public string? txt_describe_rightknee_Achy { get; set; }

        public string? cc_r_knee { get; set; }
        public string? txt_other_describe_part_Constant { get; set; }
        public string? txt_other_describe_part_Intermittent { get; set; }
        public string? txt_other_describe_part_Sharp { get; set; }
        public string? txt_other_describe_part_Electric { get; set; }
        public string? txt_other_describe_part_Shooting { get; set; }
        public string? txt_other_describe_part_Throbbing { get; set; }
        public string? txt_other_describe_part_Pulsating { get; set; }

        public string? txt_other_describe_part_Dull { get; set; }
        public string? txt_other_describe_part_Achy { get; set; }

        public string? cc_other_1 { get; set; }

        public string? other_describe_part_value { get; set; }
        public string? txt_other_describe_part { get; set; }
        public string? txt_other_describe_part_1_Constant { get; set; }
        public string? txt_other_describe_part_1_Intermittent { get; set; }
        public string? txt_other_describe_part_1_Sharp { get; set; }
        public string? txt_other_describe_part_1_Electric { get; set; }
        public string? txt_other_describe_part_1_Shooting { get; set; }
        public string? txt_other_describe_part_1_Throbbing { get; set; }
        public string? txt_other_describe_part_1_Pulsating { get; set; }
        public string? txt_other_describe_part_1_Dull { get; set; }

        public string? txt_other_describe_part_1_Achy { get; set; }
        public string? other_describe_part_1_value { get; set; }
        public string? txt_other_describe_part_1 { get; set; }

        public string? cc_other_2 { get; set; }

        //6.What relieves your pain?

        public string? relieves_pain_Resting { get; set; }
        public string? relieves_pain_Medication { get; set; }
        public string? relieves_pain_Therapy { get; set; }
        public string? relieves_pain_Sleeping { get; set; }
        public string? relieves_pain_Others { get; set; }

        public string? txt_relieves_pain { get; set; }

        //7.What increases your pain?
        public string? increase_neck_pain_lookingup { get; set; }
        public string? increase_neck_pain_lookingdown { get; set; }
        public string? increase_neck_pain_turningheadright { get; set; }
        public string? increase_neck_pain_turningheadleft { get; set; }
        public string? increase_neck_pain_driving { get; set; }
        public string? increase_neck_pain_twisting { get; set; }

        public string? increase_midback_pain_sitting { get; set; }
        public string? increase_midback_pain_standing { get; set; }
        public string? increase_midback_pain_bendingforward { get; set; }
        public string? increase_midback_pain_bendingbackwards { get; set; }
        public string? increase_midback_pain_sleeping { get; set; }

        public string? increase_midback_pain_twisting { get; set; }
        public string? increase_midback_pain_lifting { get; set; }
        public string? increase_lowback_pain_sitting { get; set; }
        public string? increase_lowback_pain_standing { get; set; }
        public string? increase_lowback_pain_bending_forward { get; set; }
        public string? increase_lowback_pain_bending_backwards { get; set; }
        public string? increase_lowback_pain_sleeping { get; set; }
        public string? increase_lowback_pain_twisting_right { get; set; }
        public string? increase_lowback_pain_twisting_left { get; set; }
        public string? increase_lowback_pain_lifting { get; set; }
        public string? increase_leftshoulder_pain_Raising_arm { get; set; }
        public string? increase_leftshoulder_pain_Lifting { get; set; }
        public string? increase_leftshoulder_pain_Working { get; set; }
        public string? increase_leftshoulder_pain_Rotation { get; set; }

        public string? increase_leftshoulder_pain_Overhead_activities { get; set; }
        public string? increase_rightshoulder_pain_Raising_arm { get; set; }
        public string? increase_rightshoulder_pain_Lifting { get; set; }
        public string? increase_rightshoulder_pain_Working { get; set; }
        public string? increase_rightshoulder_pain_Rotation { get; set; }
        public string? increase_rightshoulder_pain_Overhead_activities { get; set; }

        public string? increase_leftknee_pain_Squatting { get; set; }
        public string? increase_leftknee_pain_Walking { get; set; }
        public string? increase_leftknee_pain_Climb { get; set; }
        public string? increase_leftknee_pain_goingdown_stairs { get; set; }
        public string? increase_leftknee_pain_Standing { get; set; }

        public string? increase_leftknee_pain_getupfrom_chair { get; set; }
        public string? increase_leftknee_pain_getoutof_car { get; set; }
        public string? increase_rightknee_pain_Squatting { get; set; }
        public string? increase_rightknee_pain_Walking { get; set; }
        public string? increase_rightknee_pain_Climb { get; set; }
        public string? increase_rightknee_pain_goingdownstairs { get; set; }
        public string? increase_rightknee_pain_Standing { get; set; }
        public string? increase_rightknee_pain_getupfromchair { get; set; }

        public string? increase_rightknee_pain_getoutofcar { get; set; }

        //8.Range of Motion (For Reference - normal values in brackets)
        public string? gait_Normal { get; set; }
       
        public string? txt_gait { get; set; }
        public string? cervical_pain_right { get; set; }
        public string? cervical_pain_left { get; set; }
        public string? cervical_pain_bilateral { get; set; }
        public string? cervical_pain_paraspinal { get; set; }
        public string? cervical_pain_trapezius { get; set; }
        public string? cervical_pain_upper_cervical { get; set; }
        public string? cervical_pain_trigger_points_right { get; set; }
        public string? cervical_pain_trigger_points_left { get; set; }
        public string? cervical_pain_trigger_points_bilateral { get; set; }
        public string? cervical_pain_trigger_points_rhomboid_major { get; set; }
        public string? cervical_pain_trigger_points_rhomboid { get; set; }

        public string? cervical_pain_trigger_points_latissimus { get; set; }
        public string? cervical_pain_trigger_points_erector { get; set; }
        public string? rom_ff_fortyfive { get; set; }
        public string? rom_ext_fortyfive { get; set; }
        public string? rom_lat_bend_fortyfive_L { get; set; }
        public string? rom_lat_bend_fortyfive_R { get; set; }
        public string? rom_rot_eighty_L { get; set; }
        public string? rom_rot_eighty_R { get; set; }
        public string? rom_with_pain { get; set; }
        public string? rom_with_discomfort { get; set; }
        public string? spurlings_positive { get; set; }
        //public string? spurlings_negative { get; set; }
        public string? spurlings_bilaterally { get; set; }
        public string? spurlings_right { get; set; }
        public string? spurlings_left { get; set; }
        public string? spurlings_bilaterally_reproducing { get; set; }
        public string? txt_spurlings { get; set; }
        public string? tinels_positive { get; set; }
        // public string? tinels_negative { get; set; }
        public string? tinels_bilaterally { get; set; }
        public string? tinels_bilaterally_right { get; set; }
        public string? tinels_bilaterally_left { get; set; }
        public string? phalens_positive { get; set; }
        //public string? phalens_negative { get; set; }
        public string? phalens_bilaterally { get; set; }
        public string? phalens_bilaterally_right { get; set; }
        public string? phalens_bilaterally_left { get; set; }
        public string? facet_positive { get; set; }
        //public string? facet_negative { get; set; }
        public string? facet_bilaterally { get; set; }
        public string? facet_bilaterally_right { get; set; }
        public string? facet_bilaterally_left { get; set; }




        ////9.Neuro. Exam
        public string? dtrs_two_excepts_biceps_L { get; set; }
        // public string? dtrs_two_excepts_biceps_R { get; set; }
        public string? triceps_L { get; set; }
        // public string? triceps_R { get; set; }
        public string? brachioradialis_L { get; set; }
        public string? brachioradialis_R { get; set; }
        public string? sensory_Normal { get; set; }
        public string? sensory_Decreased { get; set; }
        public string? sensory_right { get; set; }
        public string? sensory_Left { get; set; }
        public string? sensory_C1C2 { get; set; }
        public string? sensory_C2c3 { get; set; }
        public string? sensory_c3c4 { get; set; }
        public string? sensory_c4c5 { get; set; }
        public string? sensory_c5c6 { get; set; }
        public string? sensory_c6c7 { get; set; }
        public string? upper_ext_normal { get; set; }
        public string? except_mild { get; set; }
        //public string? except_moderate { get; set; }
        //public string? except_severe { get; set; }
        public string? weakness_flexors { get; set; }
        //public string? weakness_extensors { get; set; }
        public string? thoracic_right { get; set; }
        public string? thoracic_left { get; set; }
        public string? thoracic_bilateral { get; set; }
        public string? thoracic_paraspinal_muscles { get; set; }
        public string? thoracic_trigger_points_right { get; set; }
        public string? thoracic_trigger_points_left { get; set; }
        public string? thoracic_trigger_points_bilateral { get; set; }
        public string? thoracic_trigger_points_paraspinals { get; set; }
        public string? lumbar_right { get; set; }
        public string? lumbar_left { get; set; }
        public string? lumbar_bilateral { get; set; }
        public string? lumbar_paravertebral { get; set; }
        public string? lumbar_trigger_points_right { get; set; }
        public string? lumbar_trigger_points_left { get; set; }
        public string? lumbar_trigger_points_bilateral { get; set; }
        public string? lumbar_trigger_points_quadratus { get; set; }
        public string? lumbar_trigger_points_gluteus { get; set; }
        public string? lumbar_trigger_points_lumbar { get; set; }

        public string? rom_ff_ninety { get; set; }
        public string? rom_ext_30 { get; set; }
        public string? rom_lat_bend_30_L { get; set; }
        public string? rom_lat_bend_30_R { get; set; }
        public string? neuro_rom_rot_eighty_L { get; set; }
        public string? neuro_rom_rot_eighty_R { get; set; }
        public string? neuro_rom_with_pain { get; set; }
        public string? neuro_rom_with_discomfort { get; set; }
        public string? slr_positive { get; set; }
        //public string? slr_nagative { get; set; }
        public string? slr_bilaterally { get; set; }
        public string? slr_bilaterally_right { get; set; }
        public string? slr_bilaterally_left { get; set; }
        public string? neuro_facet_positive { get; set; }
        // public string? neuro_facet_negative { get; set; }
        public string? neuro_facet_bilaterally { get; set; }
        public string? neuro_facet_bilaterally_right { get; set; }
        public string? neuro_facet_bilaterally_left { get; set; }
        public string? dtrs_two_except_pattler_L { get; set; }
        //public string? dtrs_two_except_pattler_R { get; set; }
        public string? achilles_L { get; set; }
        // public string? achilles_R { get; set; }
        public string? planter_L { get; set; }
        //public string? planter_R { get; set; }
        public string? sensory_1_Normal { get; set; }
        public string? sensory_1_Decreased { get; set; }
        public string? sensory_1_bilateral { get; set; }
        public string? sensory_1_right { get; set; }
        public string? sensory_1_Left { get; set; }
        public string? sensory_1_l1l2 { get; set; }
        public string? sensory_1_l2l3 { get; set; }
        public string? sensory_1_l3l4 { get; set; }
        public string? sensory_1_l4l5 { get; set; }
        public string? sensory_1_l5s1 { get; set; }

        public string? lower_ext_normal { get; set; }
        public string? except_1_mild { get; set; }
        //public string? except_1_moderate { get; set; }
        // public string? except_1_severe { get; set; }
        public string? weakness_1_flexors { get; set; }
        //public string? weakness_1_extensors { get; set; }
        public string? ls_swelling_positive { get; set; }
        //public string? ls_swelling_negative { get; set; }
        public string? ls_echymosis_positive { get; set; }
        //public string? ls_echymosis__negative { get; set; }
        public string? ls_deformity_positive { get; set; }
        //public string? ls_deformity_negative { get; set; }
        public string? ls_masses_positive { get; set; }
        //public string? ls_masses_negative { get; set; }
        public string? ls_tenderness_positive { get; set; }
        // public string? ls_tenderness_negative { get; set; }
        public string? ls_acjoint { get; set; }
        public string? ls_scapula { get; set; }
        public string? ls_anterior { get; set; }
        public string? ls_supraspinatus { get; set; }
        public string? ls_trapezius { get; set; }

        public string? ls_abd_oneeighty_L { get; set; }
        public string? ls_abd_oneeighty_R { get; set; }
        public string? ls_flex_oneeighty_L { get; set; }
        public string? ls_flex_oneeighty_R { get; set; }
        public string? ls_add_thirty_L { get; set; }
        public string? ls_add_thirty_R { get; set; }
        public string? ls_with_pain { get; set; }
        public string? ls_with_pain_discomfort { get; set; }
        public string? ls_ext_sixty_L { get; set; }
        public string? ls_ext_sixty_R { get; set; }
        public string? ls_ext_rot_ninety_L { get; set; }
        public string? ls_ext_rot_ninety_R { get; set; }
        public string? ls_int_rot_seventy_L { get; set; }
        public string? ls_int_rot_seventy_R { get; set; }
        public string? ls_with_pain_1 { get; set; }
        public string? ls_with_pain_1_discomfort { get; set; }
        public string? ls_hawkins_positive { get; set; }
        // public string? ls_hawkins_negative { get; set; }
        public string? ls_o_brien_test_positive { get; set; }
        // public string? ls_o_brien_test_negative { get; set; }
        public string? ls_impegement_positive { get; set; }
        //public string? ls_impegement_negative { get; set; }
        public string? ls_drop_arm_positive { get; set; }
        //public string? ls_drop_arm_negative { get; set; }
        public string? ls_cross_over_positive { get; set; }
        //public string? ls_cross_over_negative { get; set; }
        public string? ls_empty_can_positive { get; set; }
        //public string? ls_empty_can_negative { get; set; }
        public string? rs_swelling_positive { get; set; }
        //public string? rs_swelling_negative { get; set; }
        public string? rs_echymosis_positive { get; set; }
        // public string? rs_echymosis_negative { get; set; }
        public string? rs_deformity_positive { get; set; }
        // public string? rs_deformity_negative { get; set; }
        public string? rs_masses_positive { get; set; }
        // public string? rs_masses_negative { get; set; }
        public string? rs_tenderness_positive { get; set; }
        // public string? rs_tenderness_negative { get; set; }
        public string? rs_acjoint { get; set; }
        public string? rs_acjoint_scapula { get; set; }
        public string? rs_acjoint_anterior { get; set; }
        public string? rs_acjoint_supraspinatus { get; set; }
        public string? rs_acjoint_trapezius { get; set; }


        public string? rs_abd_oneeighty_L { get; set; }
        public string? rs_abd_oneeighty_R { get; set; }
        public string? rs_flex_oneeighty_L { get; set; }
        public string? rs_flex_oneeighty_R { get; set; }
        public string? rs_add_thirty_l { get; set; }
        public string? rs_add_thirty_R { get; set; }
        public string? rs_with_pain { get; set; }
        public string? rs_with_pain_discomfort { get; set; }
        public string? rs_ext_sixty_L { get; set; }
        public string? rs_ext_sixty_R { get; set; }
        public string? rs_ext_rot_ninety_L { get; set; }
        public string? rs_ext_rot_ninety_R { get; set; }
        public string? rs_int_rot_seventy_L { get; set; }
        public string? rs_int_rot_seventy_R { get; set; }
        public string? rs_with_pain_1 { get; set; }
        // public string? rs_with_pain_1_discomfort { get; set; }
        public string? rs_hawkins_positive { get; set; }
        // public string? rs_hawkins_negative { get; set; }
        public string? rs_o_brien_test_positive { get; set; }
        // public string? rs_o_brien_test_negative { get; set; }
        public string? rs_impegement_positive { get; set; }
        //  public string? rs_impegement_negative { get; set; }
        public string? rs_drop_arm_positive { get; set; }
        // public string? rs_drop_arm_negative { get; set; }
        public string? rs_cross_over_positive { get; set; }
        //public string? rs_cross_over_negative { get; set; }
        public string? rs_empty_can_positive { get; set; }
        // public string? rs_empty_can_negative { get; set; }
        public string? lk_swelling_positive { get; set; }
        // public string? lk_swelling_negative { get; set; }
        public string? lk_echymosis_positive { get; set; }
        // public string? lk_echymosis_negative { get; set; }
        public string? lk_deformity_positive { get; set; }
        //public string? lk_deformity_negative { get; set; }
        public string? lk_masses_positive { get; set; }
        // public string? lk_masses_negative { get; set; }
        public string? lk_effusion_positive { get; set; }
        // public string? lk_effusion_negative { get; set; }
        public string? lk_crepitus_positive { get; set; }
        // public string? lk_crepitus_negative { get; set; }
        public string? lk_tenderness_positive { get; set; }
        // public string? lk_tenderness_negative { get; set; }
        public string? lk_metal_joint_line { get; set; }
        public string? lk_lateral_joint_line { get; set; }
        public string? lk_patellar_tendon { get; set; }
        public string? lk_flex_oneforty_L { get; set; }
        public string? lk_flex_oneforty_R { get; set; }
        public string? lk_ext_zero_L { get; set; }
        public string? lk_ext_zero_R { get; set; }
        public string? lk_with_pain { get; set; }
        //public string? lk_with_pain_discomfort { get; set; }
        public string? lk_mcmurray_positive { get; set; }
        // public string? lk_mcmurray_negative { get; set; }
        public string? lk_anterior_positive { get; set; }
        // public string? lk_anterior_negative { get; set; }
        public string? lk_lachman_positive { get; set; }
        // public string? lk_lachman_negative { get; set; }
        public string? lk_posterior_positive { get; set; }
        //public string? lk_posterior_negative { get; set; }
        public string? lk_valgus_positive { get; set; }
        //public string? lk_valgus_negative { get; set; }
        public string? lk_varus_positive { get; set; }
        //  public string? lk_varus_negative { get; set; }
        public string? rk_swelling_positive { get; set; }
        // public string? rk_swelling_negative { get; set; }
        public string? rk_echymosis_positive { get; set; }
        // public string? rk_echymosis_negative { get; set; }
        public string? rk_deformity_positive { get; set; }
        // public string? rk_deformity_negative { get; set; }
        public string? rk_masses_positive { get; set; }
        // public string? rk_masses_negative { get; set; }
        public string? rk_effusion_positive { get; set; }
        // public string? rk_effusion_negative { get; set; }
        public string? rk_crepitus_positive { get; set; }
        //  public string? rk_crepitus_negative { get; set; }
        public string? rk_tenderness_positive { get; set; }
        // public string? rk_tenderness_negative { get; set; }
        public string? rk_metal_joint_line { get; set; }
        public string? rk_lateral_joint_line { get; set; }
        public string? rk_patellar_joint_line { get; set; }
        public string? rk_flex_oneforty_L { get; set; }
        public string? rk_flex_oneforty_R { get; set; }
        public string? rk_ext_zero_L { get; set; }
        public string? rk_ext_zero_R { get; set; }
        public string? rk_with_pain { get; set; }
        //public string? rk_with_pain_discomfort { get; set; }
        public string? rk_mcmurray_poitive { get; set; }
        // public string? rk_mcmurray_negative { get; set; }
        public string? rk_anterior_poitive { get; set; }
        // public string? rk_anterior_negative { get; set; }
        public string? rk_lachman_poitive { get; set; }
        // public string? rk_lachman_negative { get; set; }
        public string? rk_posterior_poitive { get; set; }
        // public string? rk_posterior_negative { get; set; }
        public string? rk_valgus_poitive { get; set; }
        // public string? rk_valgus_negative { get; set; }
        public string? rk_varus_poitive { get; set; }
        //public string? rk_varus_negative { get; set; }
        public string? neuro_others { get; set; }
        //10.DIAGNOSES: 
        public string? diagonosis { get; set; }

        //11.RECOMMENDATIONS: 
        public string? diag_test_refferal_emg { get; set; }
        public string? diag_test_refferal_mri { get; set; }
        public string? diag_test_refferal_xray { get; set; }
        public string? diag_test_refferal_ct { get; set; }
        public string? recommend_providerseen_Chiro { get; set; }
        public string? recommend_providerseen_Spine { get; set; }
        public string? recommend_providerseen_PainManagemant { get; set; }
        public string? recommend_providerseen_PT { get; set; }
        public string? recommend_providerseen_Nero { get; set; }
        public string? recommend_providerseen_Ortho { get; set; }
        public string? recommend_providerseen_other { get; set; }
        public string? recommend_providerseen_other_1 { get; set; }
        public string? refferal_Chiro { get; set; }
        public string? refferal_PT { get; set; }
        public string? refferal_Spine { get; set; }
        public string? refferal_Ortho { get; set; }
        public string? refferal_Psychologist { get; set; }
        public string? refferal_Physiatrist { get; set; }
        public string? refferal_NeuroLogist { get; set; }
        public string? refferal_other_1 { get; set; }

        public string? refferal_other { get; set; }
        public string? request_schedule_Discectomy { get; set; }
        public string? request_schedule_Epidurogram { get; set; }
        public string? request_schedule_Discogram { get; set; }
        public string? request_schedule_SI_Fusion { get; set; }
        public string? request_schedule_SI_Injection { get; set; }
        public string? request_schedule_SCS_Trail { get; set; }
        public string? request_schedule_SCS_Implant { get; set; }
        public string? request_schedule_MRI { get; set; }
        public string? request_schedule_EMG_NCV { get; set; }
        public string? request_schedule_Epidural { get; set; }

        public string? request_schedule_other { get; set; }
        public string? request_schedule_other_1 { get; set; }
        //12.DME Dispensed at today's office visit:
        public string? dme_dispenced_L0637 { get; set; }
        public string? dme_dispenced_E0730 { get; set; }
        public string? dme_dispenced_E0731 { get; set; }
        public string? dme_dispenced_L1832 { get; set; }
        public string? dme_dispenced_L1843 { get; set; }
        public string? dme_dispenced_L1971 { get; set; }
        public string? dme_dispenced_L1902 { get; set; }
        public string? dme_dispenced_L3807 { get; set; }
        public string? dme_dispenced_L3908 { get; set; }
        public string? dme_dispenced_L3916 { get; set; }
        public string? dme_dispenced_L0456 { get; set; }
        public string? dme_dispenced_L0480 { get; set; }
        public string? dme_dispenced_E0762 { get; set; }
        public string? Rx_Meds_Refill { get; set; }


        //13.DME Meadowbrook Medical Supplies (other provider)
        public string? dme_meadowbrook_T5001 { get; set; }
        public string? dme_meadowbrook_E0855 { get; set; }
        public string? dme_meadowbrook_E0217 { get; set; }

        //14.Procedure done today:
        public string? procedure_done_today_UTPI { get; set; }
        public string? procedure_done_today_LTPI { get; set; }
        public string? procedure_done_today_TTPI { get; set; }
        public string? procedure_done_today_Diagnostic { get; set; }
        public string? procedure_done_today_Intraarticular { get; set; }

        public string? proc_done_today { get; set; }
        //15.Followup:
        public string? followup { get; set; }


        public int? Cmp_Id { get; set; }



    }
}
