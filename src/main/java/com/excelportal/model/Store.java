package com.excelportal.model;

import javax.persistence.Column;
import javax.persistence.Entity;
import javax.persistence.GeneratedValue;
import javax.persistence.GenerationType;
import javax.persistence.Id;
import javax.persistence.JoinColumn;
import javax.persistence.ManyToOne;
import javax.persistence.Table;

@Entity
@Table(name = "stores")
public class Store {
	
	@Id
	@GeneratedValue(strategy = GenerationType.IDENTITY)
	@Column(name = "id")
	private int id;
	
	@Column(name = "store_number")
	private int storeNumber;
	
	@Column(name = "store_name")
	private String storeName;
	
	@Column(name = "city")
	private String city;
	
	@Column(name = "state")
	private String state;
	
	@Column(name = "zip")
	private int zip;
	
	@ManyToOne
	@JoinColumn(name = "coach_id")
	private AreaCoach coach;
	
	public Store() {
		
	}

	public Store(int storeNumber, String storeName, String city, String state, int zip, AreaCoach coach) {
		this.storeNumber = storeNumber;
		this.storeName = storeName;
		this.city = city;
		this.state = state;
		this.zip = zip;
		this.coach = coach;
	}

	public int getId() {
		return id;
	}

	public void setId(int id) {
		this.id = id;
	}

	public int getStoreNumber() {
		return storeNumber;
	}

	public void setStoreNumber(int storeNumber) {
		this.storeNumber = storeNumber;
	}

	public String getStoreName() {
		return storeName;
	}

	public void setStoreName(String storeName) {
		this.storeName = storeName;
	}

	public String getCity() {
		return city;
	}

	public void setCity(String city) {
		this.city = city;
	}

	public String getState() {
		return state;
	}

	public void setState(String state) {
		this.state = state;
	}

	public int getZip() {
		return zip;
	}

	public void setZip(int zip) {
		this.zip = zip;
	}

	public AreaCoach getCoach() {
		return coach;
	}

	public void setCoach(AreaCoach coach) {
		this.coach = coach;
	}

	@Override
	public String toString() {
		return "Store [id=" + id + ", storeNumber=" + storeNumber + ", storeName=" + storeName + ", city=" + city
				+ ", state=" + state + ", zip=" + zip + ", coach=" + coach + "]";
	}
	
}
