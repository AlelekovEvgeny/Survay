package nsk.tfoms.survay.pojo;

import nsk.tfoms.survay.entity.QuestionManyClinic;
import nsk.tfoms.survay.entity.SurvayClinicSec1;
import nsk.tfoms.survay.entity.SurvayClinicSec2;
import nsk.tfoms.survay.entity.SurvayClinicSec25;
import nsk.tfoms.survay.entity.SurvayClinicSecondlevel;
/*
 * ����� ��������
 * ������ ������� �������� � �������. �������� ����������� �� �� ��������
 * ����� WrapMany ��� ���������� ������� ����� ���� ��������� ������� 
 */
import nsk.tfoms.survay.entity.secondlevelDayStacionar.DayStacionarSecondlevel;
import nsk.tfoms.survay.entity.secondlevelDayStacionar.SCDSSLSec15;
import nsk.tfoms.survay.entity.secondlevelDayStacionar.SCDSSLSec2;
import nsk.tfoms.survay.entity.secondlevelDayStacionar.SCDSSLSec25;
import nsk.tfoms.survay.entity.secondlevelDayStacionar.SurvayClinicDayStacionarSec1;

public class SenderDSSL {
	
	private DayStacionarSecondlevel survay1;
	private SurvayClinicDayStacionarSec1 survay2;
	private SCDSSLSec2 survay3;
	private SCDSSLSec15 survay4;
	private SCDSSLSec25 survay5;
	public DayStacionarSecondlevel getSurvay1() {
		return survay1;
	}
	public void setSurvay1(DayStacionarSecondlevel survay1) {
		this.survay1 = survay1;
	}
	public SurvayClinicDayStacionarSec1 getSurvay2() {
		return survay2;
	}
	public void setSurvay2(SurvayClinicDayStacionarSec1 survay2) {
		this.survay2 = survay2;
	}
	public SCDSSLSec2 getSurvay3() {
		return survay3;
	}
	public void setSurvay3(SCDSSLSec2 survay3) {
		this.survay3 = survay3;
	}
	public SCDSSLSec15 getSurvay4() {
		return survay4;
	}
	public void setSurvay4(SCDSSLSec15 survay4) {
		this.survay4 = survay4;
	}
	public SCDSSLSec25 getSurvay5() {
		return survay5;
	}
	public void setSurvay5(SCDSSLSec25 survay5) {
		this.survay5 = survay5;
	}
	
		

}
