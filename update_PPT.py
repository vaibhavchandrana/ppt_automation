from pptx import Presentation   # importing the pptx library
from pptx.chart.data import ChartData
from pptx.enum.shapes import MSO_SHAPE_TYPE, MSO_AUTO_SHAPE_TYPE
from pptx.util import Pt   # importing point for setting font size
from pptx.dml.color import RGBColor  # importing RGB colors to set colors
# importing functions to create new charts
from pptx.chart.data import CategorySeriesData, CategoryChartData

# Open the existing PowerPoint presentation
try:
    ppt = Presentation('input.pptx')
    # Access the slides in the presentation
    slides = ppt.slides

# ============================ code for second slide start ================================

# Data which has to be reflected in first slide
    first_data = {
        "passerby_contribution": 15,
        "inc_in_customer": 20,
        "women_walk_in_contribution": 8,
        "cust_notice_digital_window": 3.5,
        "cust_walked_inside_digital": 14,
        "cust_walked_inside_static": 90,
        "cust_notices_light_window": 40,
        "cust_notices_static_window": 3
    }

    # checking that is data recieved is valid or not
    validation = True     # adding a validation variable to check data is valid or not

    for key, value in first_data.items():
        if type(value) == int or type(value) == float:  # checking type of values is int or float

            if value > 100:                       # if value is greater than 100 then data is wrong
                validation = False
                print(key, " has value ", value,
                      " which can not be greater than 100 ")

        else:
            validation = False
            print(key, " has value  ", value,
                  " which is not valid ( it should be digit ) ")

    # if data is valid then only perform below operation
    if validation:
        # access the second slide where modification has to be done (0-indexed)
        slide = ppt.slides[1]

        # loop for traversing all the shapes
        for shape in slide.shapes:
            # checking that shape is equal to desired shaped

            if shape.name == "Rectangle 4":
                # fetching all the paragraph inside this component
                para = shape.text_frame.paragraphs

                # changing the paragraph with updated data
                para[3].text = "Bring price point communication & Women window/easel stand – Improved passerby contribution by {}% ( {}% more customers stepped inside the store vs other stores)".format(
                    first_data['passerby_contribution'], first_data['inc_in_customer'])
                # changing the font size of current paragraph
                para[3].font.size = Pt(12)

                # changing the paragraph with updated data
                para[4].text = "Similarly, women walk-in’s contribution improved by {}% vs other stores".format(
                    first_data['women_walk_in_contribution'])
                # changing the font size of current paragraph
                para[4].font.size = Pt(12)

                # changing the paragraph with updated data
                para[10].text = "{}X customers noticed digital window vs static window".format(
                    first_data['cust_notice_digital_window'])
                # changing the font size of current paragraph
                para[10].font.size = Pt(12)

                # changing the paragraph with updated data
                para[11].text = "{}% customers walked inside store after noticing digital window Vs {}% walked in after noticing static window".format(
                    first_data['cust_walked_inside_digital'], first_data['cust_walked_inside_static'])
                # changing the font size of current paragraph
                para[11].font.size = Pt(12)

                # changing the paragraph with updated data
                para[14].text = "Only {}% customers noticed light window, whereas {}% customers notice static window".format(
                    first_data['cust_notices_light_window'], first_data['cust_notices_static_window'])
                # changing the font size of current paragraph
                para[14].font.size = Pt(12)

        # saving the modified slide
        try:
            ppt.save('output.pptx')

        except Exception as e:
            print(e)
            print("File is already opened . Please close the file first to write")

    else:
        print("Data of second slide is not valid")


# ============================ code for third slide start ================================

    # Data which has to be reflected
    data = {
        "passerby_vs_store_ff": 17.5,
        "Gender_Split_inside_store": "57:53",
        "Average_time_spent": 12.7,
        "Footfal_abandon": 18.3,
        "Dwell_10_min": 44.3,
        "Dwell_time_vs_Bill_conversion": 30.8,
        "market_footfall_walk_ins": 10,
        "percentage_of_customer_spent": 54
    }
    # checking that is data recieved is valid or not
    validation = True     # adding a validation variable to check data is valid or not

    for key, value in data.items():
        if type(value) == int or type(value) == float:

            if value > 100 and key != "Average_time_spent":
                validation = False
                print(key, " has value ", value,
                      " which can not be greater than 100 ")

        else:
            if key == "Gender_Split_inside_store":
                continue
            else:
                validation = False
                print(key, " has value  ", value,
                      " which is not valid ( it should be digit )")

    if validation:
        slide = ppt.slides[2]       # Getting the third slide to modify
        # loop for traversing all the shapes

        for shape in slide.shapes:
            # Condition for getting the desired shape( Component )

            if shape.name == "TextBox 8":
                # getting all the paragraph inside the box
                para = shape.text_frame.paragraphs

                # modifying the first paragraph
                para[0].text = "{}%".format(data['passerby_vs_store_ff'])  # Modifying the data
                para[0].font.size = Pt(25)  # Changing the font size
                para[0].font.bold = True   # making the font bold

            # Condition for getting the desired shape( Component )
            if shape.name == "TextBox 15":
                para = shape.text_frame.paragraphs
                # modifying the first paragraph
                para[0].text = "{}".format(data['Gender_Split_inside_store'])  # Modifying the data
                para[0].font.size = Pt(25)  # Changing the font size
                para[0].font.bold = True   # making the font bold

            # Condition for getting the desired shape( Component )
            if shape.name == "TextBox 22":
                para = shape.text_frame.paragraphs
                # modifying the first paragraph
                para[0].text = "{} Min".format(data['Average_time_spent'])  # Modifying the data
                para[0].font.size = Pt(25)  # Changing the font size
                para[0].font.bold = True    # making the font bold

            # Condition for getting the desired shape( Component )
            if shape.name == "TextBox 36":
                para = shape.text_frame.paragraphs
                # modifying the first paragraph
                para[0].text = "{}%".format(data['Dwell_10_min'])   # Modifying the data
                para[0].font.size = Pt(25)  # Changing the font size
                para[0].font.bold = True   # making the font bold

            # Condition for getting the desired shape( Component )
            if shape.name == "TextBox 51":
                para = shape.text_frame.paragraphs
                # modifying the first paragraph
                para[0].text = "{}%".format(data['Footfal_abandon'])  # Modifying the data
                para[0].font.size = Pt(25)  # Changing the font size
                para[0].font.bold = True      # making the font bold

            # Condition for getting the desired shape( Component )
            if shape.name == "TextBox 66":
                para = shape.text_frame.paragraphs
                # modifying the first paragraph
                para[0].text = "{}%".format(data['Dwell_time_vs_Bill_conversion'])  # Modifying the data
                para[0].font.size = Pt(25)  # Changing the font size
                para[0].font.bold = True    # making the font bold

            # Condition for getting the desired shape( Component )
            if shape.name == "Text Placeholder 3":
                # changing the title of slide
                para = shape.text_frame.paragraphs
                # modifying the first paragraph
                para[0].text = "{}% market footfall walk-ins inside the stores, Men’s walk-ins is higher Average {}% customers spent > 10 min inside the stores (potential customers)".format(
                    data['market_footfall_walk_ins'], data['percentage_of_customer_spent'])   # Modifying the data of title
                # Changing the color of text size
                para[0].font.color.rgb = RGBColor(255, 0, 0)
                para[0].font.size = Pt(22)  # Changing the font size
                para[0].font.bold = True    # making the font bold

         # saving the third slide
        try:
            ppt.save('output.pptx')  # saving the modified slide
        except Exception as e:
            print(e)
            print("File is already opened . Please close the file first to write")

    else:
        print("Data of third slide is not valid")


# ============================ code for forth slide start ================================
    # data which has to modified in chart
    new_chart_data = {
        "passerby": [410689, 366210, 372079, 371521, 373346],
        "footfall": [73920, 55713, 52936, 50434, 48136]
    }
    # validating the incoming data
    chart_data_validation = True
    # create a series for drawing line chart
    passerby_vs_footfall = [0]*len(new_chart_data['passerby'])

    for i in range(0, len(new_chart_data['passerby'])):

        # checking if there is no passerby then there should be no footfall
        if new_chart_data['passerby'][i] == 0:
            if new_chart_data['footfall'][i] == 0:
                passerby_vs_footfall[i]=0
            else:
                chart_data_validation = False
                print("passerby value and footfall value is not correct")

     
        else:
            # checking relation between footfall and passerby is feasable or not
            per=round((new_chart_data['footfall'][i]/new_chart_data['passerby'][i])*100,2)
            if per < 0.5:
                chart_data_validation = False
                print("passerby value and footfall value is not correct")
                break

            # checking if footfall data is greater than passerby data then validation failed
            if new_chart_data['footfall'][i] > new_chart_data['passerby'][i]:
                chart_data_validation = False
                print("footfall data is not valid ")
                break

            else:
                passerby_vs_footfall[i] = round((new_chart_data['footfall'][i]/new_chart_data['passerby'][i])*100)

    # if data is validated then perform below operation
    if chart_data_validation:

        slide = ppt.slides[3]        # accesing the fourth slide
        shape = slide.shapes         # fetching all the shapes in the slide

        # getting the bar chart component from the shapes
        bar_chart = slide.shapes[0].chart
        # getting the line chart component from the shapes
        line_chart = slide.shapes[1].chart

        chart_data = ChartData()      # creating new instance of chart data

        # creating a instance for new chart
        new_data = CategoryChartData()

        # copying categoring from old charts
        new_data.categories = bar_chart.plots[0].categories

        # adding series of column chart with new data
        new_series = new_data.add_series('Passer By', new_chart_data['passerby'])

        # adding series of column chart with new data
        new_series = new_data.add_series('Footfall', new_chart_data['footfall'])

        # replacing the previous chart with new chart
        bar_chart.replace_data(new_data)

        # updating the line charts
        new_line_chart = CategoryChartData()

        # copying categoring from old charts
        new_line_chart.categories = line_chart.plots[0].categories

        # adding series of line chart with new data
        new_line_series = new_line_chart.add_series('Vs Passer By', passerby_vs_footfall)

        # replacing the previous line chart with new chart
        line_chart.replace_data(new_line_chart)

        # saving the ppt into output pptx

        try:
            ppt.save('output.pptx')  # saving the modified slide
        except Exception as e:
            print(e)
            print("File is already opened . Please close the file first to write")

    else:
        print("Chart Data recieved is not valid ")


except Exception as e:
    print(e)
    print("Unable to open source file")
