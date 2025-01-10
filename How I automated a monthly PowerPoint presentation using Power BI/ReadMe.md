<h1 style="text-align: center;">How I automated a monthly PowerPoint presentation using Power BI</h1>

<div style="text-align: center; margin-top: 20px; margin-bottom: 20px;">
  <p>This article has also been published in Medium. You can read it <a href="https://medium.com/@i.doganos/how-i-automated-a-monthly-powerpoint-presentation-using-power-bi-b1c2d373598d" target="_blank">here</a>.</p>
</div>

In my work, each month, we are tasked with creating a PowerPoint presentation that highlights the changes in sales between the last month and the one before. The presentation includes more than 10 slides, each having a number of arrows (indicating whether sales have risen or fallen), values, and percentage change. Unsurprisingly, we used to lose countless working hours scrambling to get the presentation ready before the deadline.

Nevertheless, as I worked on the presentation, I realized it was exactly the same every month. The only thing that changed were the numbers and the arrows (green and facing up when sales increased, red and facing down when sales decreased). Not only that. The slides were all identical with only the product being different! Since I like to automate anything that can be automated, this was clearly a task I couldn't overlook. So I decided to create the PowerPoint, automatically, using Power BI. Doing a quick research on the web I couldn't find similar projects, so I decided to share my experience here. Of course, I won't be using the company's data nor the presentation format. Instead, I will create PowerPoint slides showcasing Spain's exports and imports over the years. Let's dive into it!

### 1. Preparing the data

The trading data come from the [Spain’s official trade database](https://datacomex.comercio.es). Since the website doesn’t allow unregistered users to download large datasets, I had to restrict my search to specific products (vehicles and its subcategories), specific countries (China, Germany, Greece, Japan,  Switzerland, Turkey, USA), and limit the timeframe between 2020 and 2023. However, the concept is the same for larger databases.

After downloading a file for imports and another for exports, I made the necessary transformations within the PowerQuery, changed the name of the countries to English, and appended the imports and exports into a single query. While I won’t go into the details of the transformation process, here’s what the final query looked like:

![Merged queries of imports and exports](https://github.com/Ioannis-D/Articles/blob/main/How%20I%20automated%20a%20monthly%20PowerPoint%20presentation%20using%20Power%C2%A0BI/Images/01.png)

As you can see, the data include the Taric code (a classification system used in international trade) but not the product description. To obtain the official descriptions, I downloaded the list of [Taric codes](https://circabc.europa.eu/ui/group/0e5f18c2-4b2f-42e9-aed4-dfe50ae1263b/library/fcf7031e-f940-4d35-be0d-757688d13756/details) and loaded it into the model to merge it with the main dataset. Here’s what the list looks like when downloaded:

![Taric codes and their description](https://github.com/Ioannis-D/Articles/blob/main/How%20I%20automated%20a%20monthly%20PowerPoint%20presentation%20using%20Power%C2%A0BI/Images/02.png)

If you look closely, the list of Taric codes is much more detailed than the codes in the main dataset. For example, the taric code *01* appears as *010000000 80* , the code* 0101* is presented as *010100000 80* , etc. Wait a minute… **there’s a clear pattern here** ! The codes in the main dataset match the official list but with the final 0s and the two numbers followed by the space removed. As Power BI supports *R* and *Python* , I used regular expressions (regex) to get rid of the unnecessary numbers:

This is the code I used:

```
= Python.Execute("# 'dataset' holds the input data for this script#(lf)import pandas#(lf)dataset.loc[:, ""Goods code""] = dataset[""Goods code""].str.replace(r'(0*\s\d{2})$', '', regex=True)",[dataset=#"Removed Columns"])
```

That would be the Python code:

```
dataset.loc[:, "Goods code"] = dataset["Goods code"].str.replace(r'(0*\s\d{2})$', '', regex=True)
```

I replaced the last 0s and the two digits followed by space with nothing. Let's break down the regular expression (0*\s\d{2})$:

I'll explain it starting at the end of the expression. The \$ indicates the end of the string, ensuring the expression only matches at the end rather than anywhere within the string. The \d represents digits, the {2} specifies exactly two (so \d{2} is exactly two digits), the \s stands for a space and the 0* matches any number of 0s. But, any number of 0s at the end of the string (remember, this is why the $ is used). The final result is:

![The taric codes and their descriptions after the manipulation](https://github.com/Ioannis-D/Articles/blob/main/How%20I%20automated%20a%20monthly%20PowerPoint%20presentation%20using%20Power%C2%A0BI/Images/03.png)

After that, the only step left to complete the query is to merge the the two datasets on the Taric code (by doing a left join). The final dataset looks like this:

![The final dataset](https://github.com/Ioannis-D/Articles/blob/main/How%20I%20automated%20a%20monthly%20PowerPoint%20presentation%20using%20Power%C2%A0BI/Images/04.png)

The presentation dynamically updates each month, comparing trade flows for the current year, month, and period with the corresponding values from the previous year, month, and period. It is wise to create a Calendar table and establish a relationship with the main dataset. I used the following DAX code to create the Calendar, ensuring it only includes the date range of the main dataset:

```
Calendar = CALENDAR(
    MIN(Trade_Vehicles_2020_2023[Date]);
    MAX(Trade_Vehicles_2020_2023[Date])
)
```

Additionally, it’s necessary to generate a date column in the main dataset, representing the first day of each month:

```
Date = DATE('Trade_Vehicles_2020_2023'[Year]; Trade_Vehicles_2020_2023[Month];1)
```

After that, the final step is to create an active relationship between the two tables:

![Date relationship between the two tables](https://github.com/Ioannis-D/Articles/blob/main/How%20I%20automated%20a%20monthly%20PowerPoint%20presentation%20using%20Power%C2%A0BI/Images/05.png)

Now that the dataset is complete, we can move on to creating measures to analyze the trade flows for the previous year, previous month, and other time periods.

### 2. Creating the Measures

As mentioned earlier, it is necessary to compare trade flows between consecutive years and months. In other words, the percentage difference can be calculated using the following equation:

> ( **∑** euros **current month** / **∑** euros **previous month** ) — 1

Filtering by month makes it straightforward to retrieve the first part of the equation. However, the second part needs to be calculated. The DAX language simplifies this process with build-in functions such as [*PREVIOUSYEAR()*](https://learn.microsoft.com/en-us/dax/previousyear-function-dax), [*PREVIOUSMONTH()*](https://learn.microsoft.com/en-us/dax/startofmonth-function-dax), and [*SAMEPERIODLASTYEAR()*](https://learn.microsoft.com/en-us/dax/sameperiodlastyear-function-dax)*, * among others.Filtering by month makes it straightforward to retrieve the first part of the equation. However, the second part needs to be calculated. The DAX language simplifies this process with build-in functions such as PREVIOUSYEAR(), PREVIOUSMONTH(), and SAMEPERIODLASTYEAR(), among others.

> I am aware that the functions mentioned above can sometimes cause issues and not work as expected. I, myself, have encountered problems with the PREVIOUSMONTH() function in other scenarios. However, in this example, it works as intended.

To retrieve the total amount of exports or imports in euros for the previous year or month, the following code can be used:

```
//Previous Year
Previous_Year_Euros = CALCULATE(
    SUM(Trade_Vehicles_2020_2023[Euros]);
    PREVIOUSYEAR('Calendar'[Date])
)
```

To compare the same period across different years (for example, January 2022 with January 2023, or even January & June 2022 with January & June 2023 ) the following measure can be created:

```
Same_Period_Last_Year_Euros = CALCULATE(
    SUM(Trade_Vehicles_2020_2023[Euros]);
    SAMEPERIODLASTYEAR('Calendar'[Date])
)
```

Just a short mention:

<p style="text-align: center; font-size: 1.2em; font-style: italic;">I prefer to have everything organized and easy to find so, I, usually, create new tables to place the measures.</p>

As you will see later on, we will also include “Variables” and “Text”, so whatever number needs to be calculated goes to the new created table “Values” (by creating the table with just one value: *Values = {1}* )

Before further proceeding, it’s always a good idea to verify that the measures return the expected results. To do this, I create separate cards for each measure and a couple of filters (one for the year and another for the month) to check if the values update correctly. Moreover, I modified the interactions between the filters and the cards so that year comparisons remain unaffected by the month filter and vice versa:

![Checking the measures created](https://github.com/Ioannis-D/Articles/blob/main/How%20I%20automated%20a%20monthly%20PowerPoint%20presentation%20using%20Power%C2%A0BI/Images/01.gif)

The measures are working! Let’s create the percentage difference between them:

```
// % difference between consecutive months
%_Dif_Month = ROUND(DIVIDE(SUM(Trade_Vehicles_2020_2023[Euros]);[Previous_Month_Euros]) - 1;2)

// % difference between consecutive years
%_Dif_Year = ROUND(DIVIDE(SUM(Trade_Vehicles_2020_2023[Euros]);[Same_Period_Last_Year_Euros]) - 1;2)

// % difference between current period and the same period the year before
%_Dif_Same_Period = ROUND(DIVIDE(SUM(Trade_Vehicles_2020_2023[Euros]);[Previous_Year_Euros]) - 1;2)
```

Testing the measures:

![Checking the measures created](https://github.com/Ioannis-D/Articles/blob/main/How%20I%20automated%20a%20monthly%20PowerPoint%20presentation%20using%20Power%C2%A0BI/Images/02.gif)

The numerical measures work. It is time to move to the text measures. Remember, the presentation is made of text, numbers, and arrows.

<p style="text-align: center; font-size: 1.2em; font-style: bold">···</p>

Before moving on, it’s important to stop and consider: What exactly does the presentation include? As mentioned earlier, it includes trade values in euros for each month (obviously), along also text and arrows. Some text remains static, but whenever the month or the year is mentioned, it has to be updated dynamically. The same applies for the arrows.

One thing at a time. Let’s start with the text and I will explain the reason that I created  a “*Variables* ” table to store specific measures/variables. Some of the text that appears includes statements:

**_Exports_** of total vehicles **_rose_** by **_x_%** -> (**_amount_€ April** vs **_amount_€ March**)
or

**_Exports_** of the category **_[category]_** is **_[amount]_€ in April**.

However, there are cases where small adjustments may need to be applied to all dynamic text. For example, imagine if presentation needs to be prepared in dollars instead of euros. Or, that instead of “*rose* ” or “*fell* ”, we want to use “*increased* ” and “*decreased* ”. If we create a specific text for each occasion, it would require manually updating all instances. Not only this would be time consuming, but also risks errors, such as uncorrected texts. By creating reusable measures that act as variables, a single update to a variable ensures that all dependent measures and text elements are updated automatically. So, **instead of creating a measure that would look like this**:

```
// This would be only one text of the whole presentation
Total_Vehicles = 
VAR flow = SELECTEDVALUE(Trade_Vehicles_2020_2023[Flow])
VAR amount = SWITCH (
        TRUE ();
        SUM(Trade_Vehicles_2020_2023[Euros]) <= 1E3; FORMAT (SUM(Trade_Vehicles_2020_2023[Euros]); "#,0.00" );
        SUM(Trade_Vehicles_2020_2023[Euros]) <= 1E6; FORMAT (SUM(Trade_Vehicles_2020_2023[Euros]); "#,0,.00 K" );
        SUM(Trade_Vehicles_2020_2023[Euros]) <= 1E9; FORMAT (SUM(Trade_Vehicles_2020_2023[Euros]); "#,0,,.00 M");
        FORMAT (SUM(Trade_Vehicles_2020_2023[Euros]); "#,0,,,.00 B")
    )
VAR previous_month_amount = 
    var amount_2 = [Previous_Month_Euros]
    return
    SWITCH (
        TRUE ();
        amount_2 <= 1E3; FORMAT (amount_2; "#,0.00" );
        amount_2 <= 1E6; FORMAT (amount_2; "#,0,.00 K" );
        amount_2 <= 1E9; FORMAT (amount_2; "#,0,,.00 M");
        FORMAT (amount_2; "#,0,,,.00 B")
    )
VAR selected_month = FORMAT(SELECTEDVALUE('Calendar'[Date]); "mmmm"; "en-US")
VAR previous_month = FORMAT(STARTOFMONTH(PREVIOUSMONTH('Calendar'[Date])); "mmmm"; "en-US")
VAR rise_fall = IF([%_Dif_Month] > 0;
    "rose";
    IF([%_Dif_Month] < 0;
    "fell";
    "remains stable"
    )
)

RETURN
COMBINEVALUES(" ";flow;"of total vehicles";rise_fall;CONCATENATE([%_Dif_Month]*100;"%");CONCATENATE("(";CONCATENATE(amount;"€"));selected_month;"vs";CONCATENATE(previous_month_amount;"€");CONCATENATE(previous_month;")"))
```

The variables are created independently and then referenced in the specific text where they are needed.

I won’t enter into technical details here, as that would make the article very long. Moreover, the purpose of writing this is to share the thought process behind automating a PowerPoint presentation using Power BI, rather than providing an in-depth explanation of DAX syntax or Power BI functionality.

<p style="text-align: center; font-size: 1.2em; font-style: italic;">To share the idea, the text is broken down into the smallest possible components, and the sentences are gradually constructed using these variables.</p>

For instance, the text of the parenthesis: **_€ [amount]_** April vs **_€ [amount]_** March is built as a variable, which is itself constructed using two other variables: one for the opening parenthesis and another for the closing parenthesis. Each one of these, n turn, relies on a variable that generates the structure **_€ [amount]_** **_([month])_**. If the presentation needs to be prepared in dollars, by simply changing the last variable, the text changes in all pages of the presentation.

The final text is stored in the “Text” table to ensure everything is organized and easy to locate and understand. Often, we create Power BI files, programs, Excel spreadsheets, etc without considering that these files might later be handed off to others within the organization. Imagine how challenging it can be for someone else to decipher the structure of a Power BI report. Even the original creator might struggle to recall the specifics after several months. **This is why organization, clear comments, and a proper documentation file are essential.**

This is how the slide of exports looks like after creating the texts:

![](https://github.com/Ioannis-D/Articles/blob/main/How%20I%20automated%20a%20monthly%20PowerPoint%20presentation%20using%20Power%C2%A0BI/Images/06.jpg)

Each element has a static filter that remains unchanged, as the presentation format is consistent each month. For instance, this slide focuses solely on exports, so the flow is filtered across the entire page. Additionally, dynamic filters, which I’ll discuss later, allow us to change the displayed values dynamically. More on this in in the following section.

Finally, let’s add some arrows. DAX supports the* UNICHAR()* function which makes it possible to access and represent a wide range of icons and special characters, including arrows. For the intermonthly changes, this is the measure of the arrow:

```
// The final measure stored in the Text table
Arrow_Month = IF([%_Dif_Month] > 0; 
                   [Arrow_Up]; 
                   IF([%_Dif_Month]<0;
                   [Arrow_Down];
                   "=")
)

// The variables stored in the Variables table
Arrow_Down = UNICHAR("129158")
Arrow_Up = UNICHAR("129157")
```

The measure is displayed on a card, and the arrow's color changes dynamically: red when exports decrease, green when exports increase, and orange when exports remain unchanged. Here's how this is achieved:

![Creating a custom function to determine the color of the arrows](https://github.com/Ioannis-D/Articles/blob/main/How%20I%20automated%20a%20monthly%20PowerPoint%20presentation%20using%20Power%C2%A0BI/Images/07.jpg)

Next, the arrows are positioned to the left of each text element, using the same filters. Finally, the slide is completed with the addition of a bar graph and some extra modifications. Here's how the final slide looks:

![Preview of the slide of exports](https://github.com/Ioannis-D/Articles/blob/main/How%20I%20automated%20a%20monthly%20PowerPoint%20presentation%20using%20Power%C2%A0BI/Images/03.gif)

To create similar slides for imports, the exports slide is duplicated, and the filter is adjusted to select imports. However, as mentioned earlier, this is not a dynamic presentation. It's a traditional PowerPoint presentation where exports, imports, and different TARIC categories each have their own static, dedicated slide. Therefore, the filters at the top of the slide need to be hidden.

This is easy to achieve, you just go to View → Selection and you hide the two filters:

![](https://github.com/Ioannis-D/Articles/blob/main/How%20I%20automated%20a%20monthly%20PowerPoint%20presentation%20using%20Power%C2%A0BI/Images/08.png)

Then, the page is duplicated, and the static filter for the flow across the entire page is updated to "Import." The filters are then synchronized between the pages. When new data for additional months or years is added, updating the dynamic filters on the first page automatically updates the entire Power BI presentation. We select from View → Sync slicers the slicers and we sync them with the rest of the pages:

![Syncing the slicers between pages](https://github.com/Ioannis-D/Articles/blob/main/How%20I%20automated%20a%20monthly%20PowerPoint%20presentation%20using%20Power%C2%A0BI/Images/09.jpg)

### 3. Transforming to PowerPoint

The presentation is ready but It is not a .ppt file (yet). This is, by far, the easiest part of the process (and the most satisfying, as all the hard work is finally done). Before exporting the pages as slides, the Power BI report needs to be published.

![Publishing the report](https://github.com/Ioannis-D/Articles/blob/main/How%20I%20automated%20a%20monthly%20PowerPoint%20presentation%20using%20Power%C2%A0BI/Images/10.png)

Once the report is opened online, there is an option to Export it as a PowerPoint presentation. This option gives two choices:

* Embed live data, or
* Export as image

In my case, I want the presentation to be static (this is the reason the filters are not shown) so I choose the second option, “*Export as image* ”. I want to export the current filtered values and all the pages that are not hidden (don’t want to include the “Test” page, of course).

![](https://github.com/Ioannis-D/Articles/blob/main/How%20I%20automated%20a%20monthly%20PowerPoint%20presentation%20using%20Power%C2%A0BI/Images/11.png)

After a short wait, a .ppt file is downloaded. The only remaining step is to remove the first page, and the presentation is ready.

Next month, the month is changed in the filter of the first page of the report, and within less than 5 minutes, the presentation is ready to be sent! A task that previously took hours or  even days of work to be prepared (imagine this with more than 10 slides, such as a slide for each country) can now be completed almost instantly and, most importantly, **without errors**.

<p style="text-align: center; font-size: 1.2em; font-style: bold">···</p>

I have always had an aversion to repetitive tasks and since I learned programming, I've always tried to automate them. While the process of automation can sometimes be time-consuming (depending on the complexity of the task and familiarity with the tools), the results are well worth the effort. Automation not only ensures accuracy and consistency but also saves valuable time, allowing us to focus on more strategic and meaningful work.

I hope this article inspires you to explore automation possibilities in your own projects and simplify your day-to-day tasks.
