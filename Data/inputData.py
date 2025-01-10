class InputData:
    def __init__(self, date, mobile, landline, first_name, last_name, position, company):
        self.date = date
        self.mobile = mobile
        self.landline = landline
        self.first_name = first_name
        self.last_name = last_name
        self.position = position
        self.company = company

    def __repr__(self):
        return f"InputData(Date={self.date}, Mobile={self.mobile}, Landline={self.landline}, First Name={self.first_name}, Last Name={self.last_name}, Position={self.position}, Company={self.company})"
