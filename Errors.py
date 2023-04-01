class TestException(Exception):
    pass

class ValidDigitError(TestException):
    def checkValidOrNot(self,key,value):
        ''' This function is used to check the  variable has valid value or not

        parameter: take a variable which has to be validated

        return : 
        
        True: if data is fine
        False : if some exception is raised
        raise exception that given digit is not a number
        '''
        validation = True
        if value > 100:                       
                    # if value is greater than 100 then data is wrong
                    validation = False
                    raise ValidDigitError(f"{key} has value {value} would not be greater than 100")

        return validation

    def checkType(self,key,value):
        ''' This function is used to check the type of a variable
        parameter: take a variable whose type has to be checked(key and value)

        return : 
        True: if data is fine
        False : if some exception is raised
        raise exception that given digit is not a number
        '''
        validation = True
        if isinstance(value, (int, float)): # checking type of values is int or float
            if  key != "Average_time_spent":
                if self.checkValidOrNot(key, value) == False:
                    validation=False
        else:
                validation = False
                raise ValidDigitError(f"{key} has the value {value} which is not valid . It should be number")
        
        return validation

        

class FeasebilityError(TestException):
    def CheckFeasibility(self,passerby,footfall):
        ''' This function is used to check the data of passerby and footfall  has valid value or not

        parameter: take two variable containing passerby and footfall value

        return : 
        True: if data is fine
        False : if some exception is raised
        
        raise exception that given digit is not a number
        '''
        chart_data_validation = True
        # checking if there is no passerby then there should be no footfall
        if passerby == 0:
            if footfall != 0:
                chart_data_validation = False
                raise FeasebilityError("passerby value and footfall value is not correct")
        
        # checking if there is footfall is greater than passerby then there is error
        if footfall > passerby:
                chart_data_validation = False
                raise FeasebilityError("passerby value and footfall value is not correct")
        
        #calculating the what is percentage of footfall with respect to passerby
        per=round((footfall/passerby)*100,2)
        # if percentage is less than 0.1 then there is some unexcepted  value 
        if per < 0.1:
            chart_data_validation = False
            raise FeasebilityError("passerby value and footfall value is not correct")

        return chart_data_validation


    
    def CheckFeasibilitySixth(self,Customer_spent,No_of_bills):
        ''' This function is used to check the data of passerby and footfall  has valid value or not

        parameter: take two variable containing passerby and footfall value

        return : 
        True: if data is fine
        False : if some exception is raised
        
        raise exception that given digit is not a number
        '''
        chart_data_validation = True
        # checking if there is no passerby then there should be no No_of_bills
        if Customer_spent == 0:
            if No_of_bills != 0:
                chart_data_validation = False
                raise FeasebilityError("Customer_spent value and No_of_bills value is not correct")
        
        # checking if there is No_of_bills is greater than passerby then there is error
        if No_of_bills > Customer_spent:
                chart_data_validation = False
                raise FeasebilityError("Customer_spent value and No_of_bills value is not correct")
        
        #calculating the what is percentage of footfall with respect to passerby
        per=round((No_of_bills/Customer_spent)*100,2)
        # if percentage is less than 0.1 then there is some unexcepted  value 
        if per < 0.1:
            chart_data_validation = False
            raise FeasebilityError("Customer_spent value and No_of_bills value is not correct")

        return chart_data_validation