# chip
dash board
{
 "cells": [
  {
   "cell_type": "markdown",
   "metadata": {},
   "source": [
    "<p style=\"text-align:center\">\n",
    "    <a href=\"https://skills.network/?utm_medium=Exinfluencer&utm_source=Exinfluencer&utm_content=000026UJ&utm_term=10006555&utm_id=NA-SkillsNetwork-Channel-SkillsNetworkCoursesIBMDA0321ENSkillsNetwork928-2022-01-01\" target=\"_blank\">\n",
    "    <img src=\"https://cf-courses-data.s3.us.cloud-object-storage.appdomain.cloud/assets/logos/SN_web_lightmode.png\" width=\"200\" alt=\"Skills Network Logo\"  />\n",
    "    </a>\n",
    "</p>\n"
   ]
  },
  {
   "cell_type": "markdown",
   "metadata": {},
   "source": [
    "# **Collecting Job Data Using APIs**\n"
   ]
  },
  {
   "cell_type": "markdown",
   "metadata": {},
   "source": [
    "Estimated time needed: **45 to 60** minutes\n"
   ]
  },
  {
   "cell_type": "markdown",
   "metadata": {},
   "source": [
    "## Objectives\n"
   ]
  },
  {
   "cell_type": "markdown",
   "metadata": {},
   "source": [
    "After completing this lab, you will be able to:\n"
   ]
  },
  {
   "cell_type": "markdown",
   "metadata": {},
   "source": [
    "*   Collect job data from Jobs API\n",
    "*   Store the collected data into an excel spreadsheet.\n"
   ]
  },
  {
   "cell_type": "markdown",
   "metadata": {},
   "source": [
    "><strong>Note: Before starting with the assignment make sure to read all the instructions and then move ahead with the coding part.</strong>\n"
   ]
  },
  {
   "cell_type": "markdown",
   "metadata": {},
   "source": [
    "#### Instructions\n"
   ]
  },
  {
   "cell_type": "markdown",
   "metadata": {},
   "source": [
    "To run the actual lab, firstly you need to click on the [Jobs_API](https://cf-courses-data.s3.us.cloud-object-storage.appdomain.cloud/IBM-DA0321EN-SkillsNetwork/labs/module%201/Accessing%20Data%20Using%20APIs/Jobs_API.ipynb) notebook link. The file contains flask code which is required to run the Jobs API data.\n",
    "\n",
    "Now, to run the code in the file that opens up follow the below steps.\n",
    "\n",
    "Step1: Download the file. \n",
    "\n",
    "Step2: Upload it on the IBM Watson studio. (If IBM Watson Cloud service does not work in your system, follow the alternate Step 2 below)\n",
    "\n",
    "Step2(alternate): Upload it in your SN labs environment using the upload button which is highlighted in red in the image below:\n",
    "Remember to upload this Jobs_API file in the same folder as your current .ipynb file\n",
    "\n",
    "<img src=\"https://cf-courses-data.s3.us.cloud-object-storage.appdomain.cloud/IBM-DA0321EN-SkillsNetwork/labs/module%201/Accessing%20Data%20Using%20APIs/Upload.PNG\">\n",
    "\n",
    "Step3:  Run all the cells of the Jobs_API file. (Even if you receive an asterik sign after running the last cell, the code works fine.)\n",
    "\n",
    "If you want to learn more about flask, which is optional, you can click on this link [here](https://cf-courses-data.s3.us.cloud-object-storage.appdomain.cloud/IBM-DA0321EN-SkillsNetwork/labs/module%201/Accessing%20Data%20Using%20APIs/FLASK_API.md.html).\n",
    "\n",
    "Once you run the flask code, you can start with your assignment.\n"
   ]
  },
  {
   "cell_type": "markdown",
   "metadata": {},
   "source": [
    "## Dataset Used in this Assignment\n",
    "\n",
    "The dataset used in this lab comes from the following source: https://www.kaggle.com/promptcloud/jobs-on-naukricom under the under a **Public Domain license**.\n",
    "\n",
    "> Note: We are using a modified subset of that dataset for the lab, so to follow the lab instructions successfully please use the dataset provided with the lab, rather than the dataset from the original source.\n",
    "\n",
    "The original dataset is a csv. We have converted the csv to json as per the requirement of the lab.\n"
   ]
  },
  {
   "cell_type": "markdown",
   "metadata": {},
   "source": [
    "## Warm-Up Exercise\n"
   ]
  },
  {
   "cell_type": "markdown",
   "metadata": {},
   "source": [
    "Before you attempt the actual lab, here is a fully solved warmup exercise that will help you to learn how to access an API.\n"
   ]
  },
  {
   "cell_type": "markdown",
   "metadata": {},
   "source": [
    "Using an API, let us find out who currently are on the International Space Station (ISS).<br> The API at [http://api.open-notify.org/astros.json](http://api.open-notify.org/astros.json?utm_medium=Exinfluencer&utm_source=Exinfluencer&utm_content=000026UJ&utm_term=10006555&utm_id=NA-SkillsNetwork-Channel-SkillsNetworkCoursesIBMDA0321ENSkillsNetwork21426264-2021-01-01&cm_mmc=Email_Newsletter-_-Developer_Ed%2BTech-_-WW_WW-_-SkillsNetwork-Courses-IBM-DA0321EN-SkillsNetwork-21426264&cm_mmca1=000026UJ&cm_mmca2=10006555&cm_mmca3=M12345678&cvosrc=email.Newsletter.M12345678&cvo_campaign=000026UJ) gives us the information of astronauts currently on ISS in json format.<br>\n",
    "You can read more about this API at [http://open-notify.org/Open-Notify-API/People-In-Space/](http://open-notify.org/Open-Notify-API/People-In-Space?utm_medium=Exinfluencer&utm_source=Exinfluencer&utm_content=000026UJ&utm_term=10006555&utm_id=NA-SkillsNetwork-Channel-SkillsNetworkCoursesIBMDA0321ENSkillsNetwork21426264-2021-01-01&cm_mmc=Email_Newsletter-_-Developer_Ed%2BTech-_-WW_WW-_-SkillsNetwork-Courses-IBM-DA0321EN-SkillsNetwork-21426264&cm_mmca1=000026UJ&cm_mmca2=10006555&cm_mmca3=M12345678&cvosrc=email.Newsletter.M12345678&cvo_campaign=000026UJ)\n"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 1,
   "metadata": {
    "tags": []
   },
   "outputs": [],
   "source": [
    "import requests # you need this module to make an API call\n",
    "import pandas as pd"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 2,
   "metadata": {
    "tags": []
   },
   "outputs": [],
   "source": [
    "api_url = \"http://api.open-notify.org/astros.json\" # this url gives use the astronaut data"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 3,
   "metadata": {
    "tags": []
   },
   "outputs": [],
   "source": [
    "response = requests.get(api_url) # Call the API using the get method and store the\n",
    "                                # output of the API call in a variable called response."
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 4,
   "metadata": {
    "tags": []
   },
   "outputs": [],
   "source": [
    "if response.ok:             # if all is response = requests.get(api_url) # Call the API using the get method and store the\n",
    "                                # output of the API call in a variable called response.well() no errors, no network timeouts)\n",
    "    data = response.json()  # store the result in json format in a variable called data\n",
    "                            # the variable data is of type dictionary."
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 5,
   "metadata": {
    "tags": []
   },
   "outputs": [
    {
     "name": "stdout",
     "output_type": "stream",
     "text": [
      "{'message': 'success', 'people': [{'name': 'Jasmin Moghbeli', 'craft': 'ISS'}, {'name': 'Andreas Mogensen', 'craft': 'ISS'}, {'name': 'Satoshi Furukawa', 'craft': 'ISS'}, {'name': 'Konstantin Borisov', 'craft': 'ISS'}, {'name': 'Oleg Kononenko', 'craft': 'ISS'}, {'name': 'Nikolai Chub', 'craft': 'ISS'}, {'name': \"Loral O'Hara\", 'craft': 'ISS'}], 'number': 7}\n"
     ]
    }
   ],
   "source": [
    "print(data)   # print the data just to check the output or for debugging"
   ]
  },
  {
   "cell_type": "markdown",
   "metadata": {},
   "source": [
    "Print the number of astronauts currently on ISS.\n"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 6,
   "metadata": {
    "tags": []
   },
   "outputs": [
    {
     "name": "stdout",
     "output_type": "stream",
     "text": [
      "7\n"
     ]
    }
   ],
   "source": [
    "print(data.get('number'))"
   ]
  },
  {
   "cell_type": "markdown",
   "metadata": {},
   "source": [
    "Print the names of the astronauts currently on ISS.\n"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 7,
   "metadata": {
    "tags": []
   },
   "outputs": [
    {
     "name": "stdout",
     "output_type": "stream",
     "text": [
      "There are 7 astronauts on ISS\n",
      "And their names are :\n",
      "Jasmin Moghbeli\n",
      "Andreas Mogensen\n",
      "Satoshi Furukawa\n",
      "Konstantin Borisov\n",
      "Oleg Kononenko\n",
      "Nikolai Chub\n",
      "Loral O'Hara\n"
     ]
    }
   ],
   "source": [
    "astronauts = data.get('people')\n",
    "print(\"There are {} astronauts on ISS\".format(len(astronauts)))\n",
    "print(\"And their names are :\")\n",
    "for astronaut in astronauts:\n",
    "    print(astronaut.get('name'))"
   ]
  },
  {
   "cell_type": "markdown",
   "metadata": {},
   "source": [
    "Hope the warmup was helpful. Good luck with your next lab!\n"
   ]
  },
  {
   "cell_type": "markdown",
   "metadata": {},
   "source": [
    "## Lab: Collect Jobs Data using Jobs API\n"
   ]
  },
  {
   "cell_type": "markdown",
   "metadata": {},
   "source": [
    "### Objective: Determine the number of jobs currently open for various technologies  and for various locations\n"
   ]
  },
  {
   "cell_type": "markdown",
   "metadata": {},
   "source": [
    "Collect the number of job postings for the following locations using the API:\n",
    "\n",
    "* Los Angeles\n",
    "* New York\n",
    "* San Francisco\n",
    "* Washington DC\n",
    "* Seattle\n",
    "* Austin\n",
    "* Detroit\n"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 8,
   "metadata": {
    "tags": []
   },
   "outputs": [],
   "source": [
    "#Import required libraries\n",
    "import pandas as pd\n",
    "import json"
   ]
  },
  {
   "cell_type": "markdown",
   "metadata": {},
   "source": [
    "#### Write a function to get the number of jobs for the Python technology.<br>\n",
    "> Note: While using the lab you need to pass the **payload** information for the **params** attribute in the form of **key** **value** pairs.\n",
    "  Refer the ungraded **rest api lab** in the course **Python for Data Science, AI & Development**  <a href=\"https://www.coursera.org/learn/python-for-applied-data-science-ai/ungradedLti/P6sW8/hands-on-lab-access-rest-apis-request-http?utm_medium=Exinfluencer&utm_source=Exinfluencer&utm_content=000026UJ&utm_term=10006555&utm_id=NA-SkillsNetwork-Channel-SkillsNetworkCoursesIBMDA0321ENSkillsNetwork928-2022-01-01\">link</a>\n",
    "  \n",
    " ##### The keys in the json are \n",
    " * Job Title\n",
    " \n",
    " * Job Experience Required\n",
    " \n",
    " * Key Skills\n",
    " \n",
    " * Role Category\n",
    " \n",
    " * Location\n",
    " \n",
    " * Functional Area\n",
    " \n",
    " * Industry\n",
    " \n",
    " * Role \n",
    " \n",
    "You can also view  the json file contents  from the following <a href = \"https://cf-courses-data.s3.us.cloud-object-storage.appdomain.cloud/IBM-DA0321EN-SkillsNetwork/labs/module%201/Accessing%20Data%20Using%20APIs/jobs.json\">json</a> URL.\n"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 12,
   "metadata": {
    "tags": []
   },
   "outputs": [],
   "source": [
    "api_url=\"http://127.0.0.1:5000/data\"\n",
    "def get_number_of_jobs_T(technology):\n",
    "    \n",
    "    #your code goes here \n",
    "    response = requests.get(api_url)\n",
    "    return technology,number_of_jobs"
   ]
  },
  {
   "cell_type": "markdown",
   "metadata": {},
   "source": [
    "Calling the function for Python and checking if it works.\n"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 13,
   "metadata": {
    "tags": []
   },
   "outputs": [
    {
     "name": "stdout",
     "output_type": "stream",
     "text": [
      "Number of jobs related to Python: Not found\n"
     ]
    }
   ],
   "source": [
    "import requests\n",
    "\n",
    "api_url = \"http://127.0.0.1:5000/data\"\n",
    "\n",
    "def get_number_of_jobs_T(technology):\n",
    "    if technology in data:\n",
    "        number_of_jobs = data[technology]\n",
    "        return technology, number_of_jobs\n",
    "    else:\n",
    "        return technology, \"Not found\"\n",
    "\n",
    "# Calling the function for Python and checking if it works.\n",
    "technology = \"Python\"\n",
    "technology, number_of_jobs = get_number_of_jobs_T(technology)\n",
    "print(f\"Number of jobs related to {technology}: {number_of_jobs}\") "
   ]
  },
  {
   "cell_type": "markdown",
   "metadata": {
    "tags": []
   },
   "source": [
    "#### Write a function to find number of jobs in US for a location of your choice\n"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 16,
   "metadata": {
    "tags": []
   },
   "outputs": [],
   "source": [
    "\n",
    "import requests\n",
    "\n",
    "api_url = \"http://127.0.0.1:5000/data\"\n",
    "\n",
    "def get_number_of_jobs_for_location(technology, location):\n",
    "    response = requests.get(api_url)\n",
    "    \n",
    "    # Check if the request was successful\n",
    "    if response.status_code == 200:\n",
    "        # Parse JSON response\n",
    "        data = response.json()\n",
    "        # Extract the number of jobs for the specified location and technology\n",
    "        if location in data and technology in data[location]:\n",
    "            return data[location][technology]\n",
    "        else:\n",
    "            return \"Location or technology not found in data.\"\n",
    "    else:\n",
    "        # Print error message if the request was not successful\n",
    "        print(\"Error:\", response.status_code)\n",
    "        return None\n",
    "    "
   ]
  },
  {
   "cell_type": "markdown",
   "metadata": {},
   "source": [
    "Call the function for Los Angeles and check if it is working.\n",
    "\n",
    "\n"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 17,
   "metadata": {
    "tags": []
   },
   "outputs": [
    {
     "ename": "ConnectionError",
     "evalue": "HTTPConnectionPool(host='127.0.0.1', port=5000): Max retries exceeded with url: /data (Caused by NewConnectionError('<urllib3.connection.HTTPConnection object at 0x7f1aa3f4b910>: Failed to establish a new connection: [Errno 111] Connection refused'))",
     "output_type": "error",
     "traceback": [
      "\u001b[0;31m---------------------------------------------------------------------------\u001b[0m",
      "\u001b[0;31mConnectionRefusedError\u001b[0m                    Traceback (most recent call last)",
      "\u001b[0;32m~/conda/envs/python/lib/python3.7/site-packages/urllib3/connection.py\u001b[0m in \u001b[0;36m_new_conn\u001b[0;34m(self)\u001b[0m\n\u001b[1;32m    174\u001b[0m             conn = connection.create_connection(\n\u001b[0;32m--> 175\u001b[0;31m                 \u001b[0;34m(\u001b[0m\u001b[0mself\u001b[0m\u001b[0;34m.\u001b[0m\u001b[0m_dns_host\u001b[0m\u001b[0;34m,\u001b[0m \u001b[0mself\u001b[0m\u001b[0;34m.\u001b[0m\u001b[0mport\u001b[0m\u001b[0;34m)\u001b[0m\u001b[0;34m,\u001b[0m \u001b[0mself\u001b[0m\u001b[0;34m.\u001b[0m\u001b[0mtimeout\u001b[0m\u001b[0;34m,\u001b[0m \u001b[0;34m**\u001b[0m\u001b[0mextra_kw\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0m\n\u001b[0m\u001b[1;32m    176\u001b[0m             )\n",
      "\u001b[0;32m~/conda/envs/python/lib/python3.7/site-packages/urllib3/util/connection.py\u001b[0m in \u001b[0;36mcreate_connection\u001b[0;34m(address, timeout, source_address, socket_options)\u001b[0m\n\u001b[1;32m     94\u001b[0m     \u001b[0;32mif\u001b[0m \u001b[0merr\u001b[0m \u001b[0;32mis\u001b[0m \u001b[0;32mnot\u001b[0m \u001b[0;32mNone\u001b[0m\u001b[0;34m:\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0m\n\u001b[0;32m---> 95\u001b[0;31m         \u001b[0;32mraise\u001b[0m \u001b[0merr\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0m\n\u001b[0m\u001b[1;32m     96\u001b[0m \u001b[0;34m\u001b[0m\u001b[0m\n",
      "\u001b[0;32m~/conda/envs/python/lib/python3.7/site-packages/urllib3/util/connection.py\u001b[0m in \u001b[0;36mcreate_connection\u001b[0;34m(address, timeout, source_address, socket_options)\u001b[0m\n\u001b[1;32m     84\u001b[0m                 \u001b[0msock\u001b[0m\u001b[0;34m.\u001b[0m\u001b[0mbind\u001b[0m\u001b[0;34m(\u001b[0m\u001b[0msource_address\u001b[0m\u001b[0;34m)\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0m\n\u001b[0;32m---> 85\u001b[0;31m             \u001b[0msock\u001b[0m\u001b[0;34m.\u001b[0m\u001b[0mconnect\u001b[0m\u001b[0;34m(\u001b[0m\u001b[0msa\u001b[0m\u001b[0;34m)\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0m\n\u001b[0m\u001b[1;32m     86\u001b[0m             \u001b[0;32mreturn\u001b[0m \u001b[0msock\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0m\n",
      "\u001b[0;31mConnectionRefusedError\u001b[0m: [Errno 111] Connection refused",
      "\nDuring handling of the above exception, another exception occurred:\n",
      "\u001b[0;31mNewConnectionError\u001b[0m                        Traceback (most recent call last)",
      "\u001b[0;32m~/conda/envs/python/lib/python3.7/site-packages/urllib3/connectionpool.py\u001b[0m in \u001b[0;36murlopen\u001b[0;34m(self, method, url, body, headers, retries, redirect, assert_same_host, timeout, pool_timeout, release_conn, chunked, body_pos, **response_kw)\u001b[0m\n\u001b[1;32m    709\u001b[0m                 \u001b[0mheaders\u001b[0m\u001b[0;34m=\u001b[0m\u001b[0mheaders\u001b[0m\u001b[0;34m,\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0m\n\u001b[0;32m--> 710\u001b[0;31m                 \u001b[0mchunked\u001b[0m\u001b[0;34m=\u001b[0m\u001b[0mchunked\u001b[0m\u001b[0;34m,\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0m\n\u001b[0m\u001b[1;32m    711\u001b[0m             )\n",
      "\u001b[0;32m~/conda/envs/python/lib/python3.7/site-packages/urllib3/connectionpool.py\u001b[0m in \u001b[0;36m_make_request\u001b[0;34m(self, conn, method, url, timeout, chunked, **httplib_request_kw)\u001b[0m\n\u001b[1;32m    397\u001b[0m             \u001b[0;32melse\u001b[0m\u001b[0;34m:\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0m\n\u001b[0;32m--> 398\u001b[0;31m                 \u001b[0mconn\u001b[0m\u001b[0;34m.\u001b[0m\u001b[0mrequest\u001b[0m\u001b[0;34m(\u001b[0m\u001b[0mmethod\u001b[0m\u001b[0;34m,\u001b[0m \u001b[0murl\u001b[0m\u001b[0;34m,\u001b[0m \u001b[0;34m**\u001b[0m\u001b[0mhttplib_request_kw\u001b[0m\u001b[0;34m)\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0m\n\u001b[0m\u001b[1;32m    399\u001b[0m \u001b[0;34m\u001b[0m\u001b[0m\n",
      "\u001b[0;32m~/conda/envs/python/lib/python3.7/site-packages/urllib3/connection.py\u001b[0m in \u001b[0;36mrequest\u001b[0;34m(self, method, url, body, headers)\u001b[0m\n\u001b[1;32m    243\u001b[0m             \u001b[0mheaders\u001b[0m\u001b[0;34m[\u001b[0m\u001b[0;34m\"User-Agent\"\u001b[0m\u001b[0;34m]\u001b[0m \u001b[0;34m=\u001b[0m \u001b[0m_get_default_user_agent\u001b[0m\u001b[0;34m(\u001b[0m\u001b[0;34m)\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0m\n\u001b[0;32m--> 244\u001b[0;31m         \u001b[0msuper\u001b[0m\u001b[0;34m(\u001b[0m\u001b[0mHTTPConnection\u001b[0m\u001b[0;34m,\u001b[0m \u001b[0mself\u001b[0m\u001b[0;34m)\u001b[0m\u001b[0;34m.\u001b[0m\u001b[0mrequest\u001b[0m\u001b[0;34m(\u001b[0m\u001b[0mmethod\u001b[0m\u001b[0;34m,\u001b[0m \u001b[0murl\u001b[0m\u001b[0;34m,\u001b[0m \u001b[0mbody\u001b[0m\u001b[0;34m=\u001b[0m\u001b[0mbody\u001b[0m\u001b[0;34m,\u001b[0m \u001b[0mheaders\u001b[0m\u001b[0;34m=\u001b[0m\u001b[0mheaders\u001b[0m\u001b[0;34m)\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0m\n\u001b[0m\u001b[1;32m    245\u001b[0m \u001b[0;34m\u001b[0m\u001b[0m\n",
      "\u001b[0;32m~/conda/envs/python/lib/python3.7/http/client.py\u001b[0m in \u001b[0;36mrequest\u001b[0;34m(self, method, url, body, headers, encode_chunked)\u001b[0m\n\u001b[1;32m   1280\u001b[0m         \u001b[0;34m\"\"\"Send a complete request to the server.\"\"\"\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0m\n\u001b[0;32m-> 1281\u001b[0;31m         \u001b[0mself\u001b[0m\u001b[0;34m.\u001b[0m\u001b[0m_send_request\u001b[0m\u001b[0;34m(\u001b[0m\u001b[0mmethod\u001b[0m\u001b[0;34m,\u001b[0m \u001b[0murl\u001b[0m\u001b[0;34m,\u001b[0m \u001b[0mbody\u001b[0m\u001b[0;34m,\u001b[0m \u001b[0mheaders\u001b[0m\u001b[0;34m,\u001b[0m \u001b[0mencode_chunked\u001b[0m\u001b[0;34m)\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0m\n\u001b[0m\u001b[1;32m   1282\u001b[0m \u001b[0;34m\u001b[0m\u001b[0m\n",
      "\u001b[0;32m~/conda/envs/python/lib/python3.7/http/client.py\u001b[0m in \u001b[0;36m_send_request\u001b[0;34m(self, method, url, body, headers, encode_chunked)\u001b[0m\n\u001b[1;32m   1326\u001b[0m             \u001b[0mbody\u001b[0m \u001b[0;34m=\u001b[0m \u001b[0m_encode\u001b[0m\u001b[0;34m(\u001b[0m\u001b[0mbody\u001b[0m\u001b[0;34m,\u001b[0m \u001b[0;34m'body'\u001b[0m\u001b[0;34m)\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0m\n\u001b[0;32m-> 1327\u001b[0;31m         \u001b[0mself\u001b[0m\u001b[0;34m.\u001b[0m\u001b[0mendheaders\u001b[0m\u001b[0;34m(\u001b[0m\u001b[0mbody\u001b[0m\u001b[0;34m,\u001b[0m \u001b[0mencode_chunked\u001b[0m\u001b[0;34m=\u001b[0m\u001b[0mencode_chunked\u001b[0m\u001b[0;34m)\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0m\n\u001b[0m\u001b[1;32m   1328\u001b[0m \u001b[0;34m\u001b[0m\u001b[0m\n",
      "\u001b[0;32m~/conda/envs/python/lib/python3.7/http/client.py\u001b[0m in \u001b[0;36mendheaders\u001b[0;34m(self, message_body, encode_chunked)\u001b[0m\n\u001b[1;32m   1275\u001b[0m             \u001b[0;32mraise\u001b[0m \u001b[0mCannotSendHeader\u001b[0m\u001b[0;34m(\u001b[0m\u001b[0;34m)\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0m\n\u001b[0;32m-> 1276\u001b[0;31m         \u001b[0mself\u001b[0m\u001b[0;34m.\u001b[0m\u001b[0m_send_output\u001b[0m\u001b[0;34m(\u001b[0m\u001b[0mmessage_body\u001b[0m\u001b[0;34m,\u001b[0m \u001b[0mencode_chunked\u001b[0m\u001b[0;34m=\u001b[0m\u001b[0mencode_chunked\u001b[0m\u001b[0;34m)\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0m\n\u001b[0m\u001b[1;32m   1277\u001b[0m \u001b[0;34m\u001b[0m\u001b[0m\n",
      "\u001b[0;32m~/conda/envs/python/lib/python3.7/http/client.py\u001b[0m in \u001b[0;36m_send_output\u001b[0;34m(self, message_body, encode_chunked)\u001b[0m\n\u001b[1;32m   1035\u001b[0m         \u001b[0;32mdel\u001b[0m \u001b[0mself\u001b[0m\u001b[0;34m.\u001b[0m\u001b[0m_buffer\u001b[0m\u001b[0;34m[\u001b[0m\u001b[0;34m:\u001b[0m\u001b[0;34m]\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0m\n\u001b[0;32m-> 1036\u001b[0;31m         \u001b[0mself\u001b[0m\u001b[0;34m.\u001b[0m\u001b[0msend\u001b[0m\u001b[0;34m(\u001b[0m\u001b[0mmsg\u001b[0m\u001b[0;34m)\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0m\n\u001b[0m\u001b[1;32m   1037\u001b[0m \u001b[0;34m\u001b[0m\u001b[0m\n",
      "\u001b[0;32m~/conda/envs/python/lib/python3.7/http/client.py\u001b[0m in \u001b[0;36msend\u001b[0;34m(self, data)\u001b[0m\n\u001b[1;32m    975\u001b[0m             \u001b[0;32mif\u001b[0m \u001b[0mself\u001b[0m\u001b[0;34m.\u001b[0m\u001b[0mauto_open\u001b[0m\u001b[0;34m:\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0m\n\u001b[0;32m--> 976\u001b[0;31m                 \u001b[0mself\u001b[0m\u001b[0;34m.\u001b[0m\u001b[0mconnect\u001b[0m\u001b[0;34m(\u001b[0m\u001b[0;34m)\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0m\n\u001b[0m\u001b[1;32m    977\u001b[0m             \u001b[0;32melse\u001b[0m\u001b[0;34m:\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0m\n",
      "\u001b[0;32m~/conda/envs/python/lib/python3.7/site-packages/urllib3/connection.py\u001b[0m in \u001b[0;36mconnect\u001b[0;34m(self)\u001b[0m\n\u001b[1;32m    204\u001b[0m     \u001b[0;32mdef\u001b[0m \u001b[0mconnect\u001b[0m\u001b[0;34m(\u001b[0m\u001b[0mself\u001b[0m\u001b[0;34m)\u001b[0m\u001b[0;34m:\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0m\n\u001b[0;32m--> 205\u001b[0;31m         \u001b[0mconn\u001b[0m \u001b[0;34m=\u001b[0m \u001b[0mself\u001b[0m\u001b[0;34m.\u001b[0m\u001b[0m_new_conn\u001b[0m\u001b[0;34m(\u001b[0m\u001b[0;34m)\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0m\n\u001b[0m\u001b[1;32m    206\u001b[0m         \u001b[0mself\u001b[0m\u001b[0;34m.\u001b[0m\u001b[0m_prepare_conn\u001b[0m\u001b[0;34m(\u001b[0m\u001b[0mconn\u001b[0m\u001b[0;34m)\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0m\n",
      "\u001b[0;32m~/conda/envs/python/lib/python3.7/site-packages/urllib3/connection.py\u001b[0m in \u001b[0;36m_new_conn\u001b[0;34m(self)\u001b[0m\n\u001b[1;32m    186\u001b[0m             raise NewConnectionError(\n\u001b[0;32m--> 187\u001b[0;31m                 \u001b[0mself\u001b[0m\u001b[0;34m,\u001b[0m \u001b[0;34m\"Failed to establish a new connection: %s\"\u001b[0m \u001b[0;34m%\u001b[0m \u001b[0me\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0m\n\u001b[0m\u001b[1;32m    188\u001b[0m             )\n",
      "\u001b[0;31mNewConnectionError\u001b[0m: <urllib3.connection.HTTPConnection object at 0x7f1aa3f4b910>: Failed to establish a new connection: [Errno 111] Connection refused",
      "\nDuring handling of the above exception, another exception occurred:\n",
      "\u001b[0;31mMaxRetryError\u001b[0m                             Traceback (most recent call last)",
      "\u001b[0;32m~/conda/envs/python/lib/python3.7/site-packages/requests/adapters.py\u001b[0m in \u001b[0;36msend\u001b[0;34m(self, request, stream, timeout, verify, cert, proxies)\u001b[0m\n\u001b[1;32m    497\u001b[0m                 \u001b[0mtimeout\u001b[0m\u001b[0;34m=\u001b[0m\u001b[0mtimeout\u001b[0m\u001b[0;34m,\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0m\n\u001b[0;32m--> 498\u001b[0;31m                 \u001b[0mchunked\u001b[0m\u001b[0;34m=\u001b[0m\u001b[0mchunked\u001b[0m\u001b[0;34m,\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0m\n\u001b[0m\u001b[1;32m    499\u001b[0m             )\n",
      "\u001b[0;32m~/conda/envs/python/lib/python3.7/site-packages/urllib3/connectionpool.py\u001b[0m in \u001b[0;36murlopen\u001b[0;34m(self, method, url, body, headers, retries, redirect, assert_same_host, timeout, pool_timeout, release_conn, chunked, body_pos, **response_kw)\u001b[0m\n\u001b[1;32m    787\u001b[0m             retries = retries.increment(\n\u001b[0;32m--> 788\u001b[0;31m                 \u001b[0mmethod\u001b[0m\u001b[0;34m,\u001b[0m \u001b[0murl\u001b[0m\u001b[0;34m,\u001b[0m \u001b[0merror\u001b[0m\u001b[0;34m=\u001b[0m\u001b[0me\u001b[0m\u001b[0;34m,\u001b[0m \u001b[0m_pool\u001b[0m\u001b[0;34m=\u001b[0m\u001b[0mself\u001b[0m\u001b[0;34m,\u001b[0m \u001b[0m_stacktrace\u001b[0m\u001b[0;34m=\u001b[0m\u001b[0msys\u001b[0m\u001b[0;34m.\u001b[0m\u001b[0mexc_info\u001b[0m\u001b[0;34m(\u001b[0m\u001b[0;34m)\u001b[0m\u001b[0;34m[\u001b[0m\u001b[0;36m2\u001b[0m\u001b[0;34m]\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0m\n\u001b[0m\u001b[1;32m    789\u001b[0m             )\n",
      "\u001b[0;32m~/conda/envs/python/lib/python3.7/site-packages/urllib3/util/retry.py\u001b[0m in \u001b[0;36mincrement\u001b[0;34m(self, method, url, response, error, _pool, _stacktrace)\u001b[0m\n\u001b[1;32m    591\u001b[0m         \u001b[0;32mif\u001b[0m \u001b[0mnew_retry\u001b[0m\u001b[0;34m.\u001b[0m\u001b[0mis_exhausted\u001b[0m\u001b[0;34m(\u001b[0m\u001b[0;34m)\u001b[0m\u001b[0;34m:\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0m\n\u001b[0;32m--> 592\u001b[0;31m             \u001b[0;32mraise\u001b[0m \u001b[0mMaxRetryError\u001b[0m\u001b[0;34m(\u001b[0m\u001b[0m_pool\u001b[0m\u001b[0;34m,\u001b[0m \u001b[0murl\u001b[0m\u001b[0;34m,\u001b[0m \u001b[0merror\u001b[0m \u001b[0;32mor\u001b[0m \u001b[0mResponseError\u001b[0m\u001b[0;34m(\u001b[0m\u001b[0mcause\u001b[0m\u001b[0;34m)\u001b[0m\u001b[0;34m)\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0m\n\u001b[0m\u001b[1;32m    593\u001b[0m \u001b[0;34m\u001b[0m\u001b[0m\n",
      "\u001b[0;31mMaxRetryError\u001b[0m: HTTPConnectionPool(host='127.0.0.1', port=5000): Max retries exceeded with url: /data (Caused by NewConnectionError('<urllib3.connection.HTTPConnection object at 0x7f1aa3f4b910>: Failed to establish a new connection: [Errno 111] Connection refused'))",
      "\nDuring handling of the above exception, another exception occurred:\n",
      "\u001b[0;31mConnectionError\u001b[0m                           Traceback (most recent call last)",
      "\u001b[0;32m/tmp/ipykernel_659/2001643928.py\u001b[0m in \u001b[0;36m<module>\u001b[0;34m\u001b[0m\n\u001b[1;32m      3\u001b[0m \u001b[0mlocation\u001b[0m \u001b[0;34m=\u001b[0m \u001b[0;34m'Los Angeles'\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0m\n\u001b[1;32m      4\u001b[0m \u001b[0;34m\u001b[0m\u001b[0m\n\u001b[0;32m----> 5\u001b[0;31m \u001b[0mprint\u001b[0m\u001b[0;34m(\u001b[0m\u001b[0mget_number_of_jobs_for_location\u001b[0m\u001b[0;34m(\u001b[0m\u001b[0mtechnology\u001b[0m\u001b[0;34m,\u001b[0m \u001b[0mlocation\u001b[0m\u001b[0;34m)\u001b[0m\u001b[0;34m)\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0m\n\u001b[0m\u001b[1;32m      6\u001b[0m \u001b[0;34m\u001b[0m\u001b[0m\n",
      "\u001b[0;32m/tmp/ipykernel_659/1769506934.py\u001b[0m in \u001b[0;36mget_number_of_jobs_for_location\u001b[0;34m(technology, location)\u001b[0m\n\u001b[1;32m      4\u001b[0m \u001b[0;34m\u001b[0m\u001b[0m\n\u001b[1;32m      5\u001b[0m \u001b[0;32mdef\u001b[0m \u001b[0mget_number_of_jobs_for_location\u001b[0m\u001b[0;34m(\u001b[0m\u001b[0mtechnology\u001b[0m\u001b[0;34m,\u001b[0m \u001b[0mlocation\u001b[0m\u001b[0;34m)\u001b[0m\u001b[0;34m:\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0m\n\u001b[0;32m----> 6\u001b[0;31m     \u001b[0mresponse\u001b[0m \u001b[0;34m=\u001b[0m \u001b[0mrequests\u001b[0m\u001b[0;34m.\u001b[0m\u001b[0mget\u001b[0m\u001b[0;34m(\u001b[0m\u001b[0mapi_url\u001b[0m\u001b[0;34m)\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0m\n\u001b[0m\u001b[1;32m      7\u001b[0m \u001b[0;34m\u001b[0m\u001b[0m\n\u001b[1;32m      8\u001b[0m     \u001b[0;31m# Check if the request was successful\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0m\n",
      "\u001b[0;32m~/conda/envs/python/lib/python3.7/site-packages/requests/api.py\u001b[0m in \u001b[0;36mget\u001b[0;34m(url, params, **kwargs)\u001b[0m\n\u001b[1;32m     71\u001b[0m     \"\"\"\n\u001b[1;32m     72\u001b[0m \u001b[0;34m\u001b[0m\u001b[0m\n\u001b[0;32m---> 73\u001b[0;31m     \u001b[0;32mreturn\u001b[0m \u001b[0mrequest\u001b[0m\u001b[0;34m(\u001b[0m\u001b[0;34m\"get\"\u001b[0m\u001b[0;34m,\u001b[0m \u001b[0murl\u001b[0m\u001b[0;34m,\u001b[0m \u001b[0mparams\u001b[0m\u001b[0;34m=\u001b[0m\u001b[0mparams\u001b[0m\u001b[0;34m,\u001b[0m \u001b[0;34m**\u001b[0m\u001b[0mkwargs\u001b[0m\u001b[0;34m)\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0m\n\u001b[0m\u001b[1;32m     74\u001b[0m \u001b[0;34m\u001b[0m\u001b[0m\n\u001b[1;32m     75\u001b[0m \u001b[0;34m\u001b[0m\u001b[0m\n",
      "\u001b[0;32m~/conda/envs/python/lib/python3.7/site-packages/requests/api.py\u001b[0m in \u001b[0;36mrequest\u001b[0;34m(method, url, **kwargs)\u001b[0m\n\u001b[1;32m     57\u001b[0m     \u001b[0;31m# cases, and look like a memory leak in others.\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0m\n\u001b[1;32m     58\u001b[0m     \u001b[0;32mwith\u001b[0m \u001b[0msessions\u001b[0m\u001b[0;34m.\u001b[0m\u001b[0mSession\u001b[0m\u001b[0;34m(\u001b[0m\u001b[0;34m)\u001b[0m \u001b[0;32mas\u001b[0m \u001b[0msession\u001b[0m\u001b[0;34m:\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0m\n\u001b[0;32m---> 59\u001b[0;31m         \u001b[0;32mreturn\u001b[0m \u001b[0msession\u001b[0m\u001b[0;34m.\u001b[0m\u001b[0mrequest\u001b[0m\u001b[0;34m(\u001b[0m\u001b[0mmethod\u001b[0m\u001b[0;34m=\u001b[0m\u001b[0mmethod\u001b[0m\u001b[0;34m,\u001b[0m \u001b[0murl\u001b[0m\u001b[0;34m=\u001b[0m\u001b[0murl\u001b[0m\u001b[0;34m,\u001b[0m \u001b[0;34m**\u001b[0m\u001b[0mkwargs\u001b[0m\u001b[0;34m)\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0m\n\u001b[0m\u001b[1;32m     60\u001b[0m \u001b[0;34m\u001b[0m\u001b[0m\n\u001b[1;32m     61\u001b[0m \u001b[0;34m\u001b[0m\u001b[0m\n",
      "\u001b[0;32m~/conda/envs/python/lib/python3.7/site-packages/requests/sessions.py\u001b[0m in \u001b[0;36mrequest\u001b[0;34m(self, method, url, params, data, headers, cookies, files, auth, timeout, allow_redirects, proxies, hooks, stream, verify, cert, json)\u001b[0m\n\u001b[1;32m    585\u001b[0m         }\n\u001b[1;32m    586\u001b[0m         \u001b[0msend_kwargs\u001b[0m\u001b[0;34m.\u001b[0m\u001b[0mupdate\u001b[0m\u001b[0;34m(\u001b[0m\u001b[0msettings\u001b[0m\u001b[0;34m)\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0m\n\u001b[0;32m--> 587\u001b[0;31m         \u001b[0mresp\u001b[0m \u001b[0;34m=\u001b[0m \u001b[0mself\u001b[0m\u001b[0;34m.\u001b[0m\u001b[0msend\u001b[0m\u001b[0;34m(\u001b[0m\u001b[0mprep\u001b[0m\u001b[0;34m,\u001b[0m \u001b[0;34m**\u001b[0m\u001b[0msend_kwargs\u001b[0m\u001b[0;34m)\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0m\n\u001b[0m\u001b[1;32m    588\u001b[0m \u001b[0;34m\u001b[0m\u001b[0m\n\u001b[1;32m    589\u001b[0m         \u001b[0;32mreturn\u001b[0m \u001b[0mresp\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0m\n",
      "\u001b[0;32m~/conda/envs/python/lib/python3.7/site-packages/requests/sessions.py\u001b[0m in \u001b[0;36msend\u001b[0;34m(self, request, **kwargs)\u001b[0m\n\u001b[1;32m    699\u001b[0m \u001b[0;34m\u001b[0m\u001b[0m\n\u001b[1;32m    700\u001b[0m         \u001b[0;31m# Send the request\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0m\n\u001b[0;32m--> 701\u001b[0;31m         \u001b[0mr\u001b[0m \u001b[0;34m=\u001b[0m \u001b[0madapter\u001b[0m\u001b[0;34m.\u001b[0m\u001b[0msend\u001b[0m\u001b[0;34m(\u001b[0m\u001b[0mrequest\u001b[0m\u001b[0;34m,\u001b[0m \u001b[0;34m**\u001b[0m\u001b[0mkwargs\u001b[0m\u001b[0;34m)\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0m\n\u001b[0m\u001b[1;32m    702\u001b[0m \u001b[0;34m\u001b[0m\u001b[0m\n\u001b[1;32m    703\u001b[0m         \u001b[0;31m# Total elapsed time of the request (approximately)\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0m\n",
      "\u001b[0;32m~/conda/envs/python/lib/python3.7/site-packages/requests/adapters.py\u001b[0m in \u001b[0;36msend\u001b[0;34m(self, request, stream, timeout, verify, cert, proxies)\u001b[0m\n\u001b[1;32m    518\u001b[0m                 \u001b[0;32mraise\u001b[0m \u001b[0mSSLError\u001b[0m\u001b[0;34m(\u001b[0m\u001b[0me\u001b[0m\u001b[0;34m,\u001b[0m \u001b[0mrequest\u001b[0m\u001b[0;34m=\u001b[0m\u001b[0mrequest\u001b[0m\u001b[0;34m)\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0m\n\u001b[1;32m    519\u001b[0m \u001b[0;34m\u001b[0m\u001b[0m\n\u001b[0;32m--> 520\u001b[0;31m             \u001b[0;32mraise\u001b[0m \u001b[0mConnectionError\u001b[0m\u001b[0;34m(\u001b[0m\u001b[0me\u001b[0m\u001b[0;34m,\u001b[0m \u001b[0mrequest\u001b[0m\u001b[0;34m=\u001b[0m\u001b[0mrequest\u001b[0m\u001b[0;34m)\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0m\n\u001b[0m\u001b[1;32m    521\u001b[0m \u001b[0;34m\u001b[0m\u001b[0m\n\u001b[1;32m    522\u001b[0m         \u001b[0;32mexcept\u001b[0m \u001b[0mClosedPoolError\u001b[0m \u001b[0;32mas\u001b[0m \u001b[0me\u001b[0m\u001b[0;34m:\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0m\n",
      "\u001b[0;31mConnectionError\u001b[0m: HTTPConnectionPool(host='127.0.0.1', port=5000): Max retries exceeded with url: /data (Caused by NewConnectionError('<urllib3.connection.HTTPConnection object at 0x7f1aa3f4b910>: Failed to establish a new connection: [Errno 111] Connection refused'))"
     ]
    }
   ],
   "source": [
    "#your code goes here\n",
    "technology = 'Python'\n",
    "location = 'Los Angeles'\n",
    "\n",
    "print(get_number_of_jobs_for_location(technology, location))\n",
    "\n"
   ]
  },
  {
   "cell_type": "markdown",
   "metadata": {},
   "source": [
    "### Store the results in an excel file\n"
   ]
  },
  {
   "cell_type": "markdown",
   "metadata": {},
   "source": [
    "Call the API for all the given technologies above and write the results in an excel spreadsheet.\n"
   ]
  },
  {
   "cell_type": "markdown",
   "metadata": {},
   "source": [
    "If you do not know how create excel file using python, double click here for **hints**.\n",
    "\n",
    "<!--\n",
    "\n",
    "from openpyxl import Workbook        # import Workbook class from module openpyxl\n",
    "wb=Workbook()                        # create a workbook object\n",
    "ws=wb.active                         # use the active worksheet\n",
    "ws.append(['Country','Continent'])   # add a row with two columns 'Country' and 'Continent'\n",
    "ws.append(['Eygpt','Africa'])        # add a row with two columns 'Egypt' and 'Africa'\n",
    "ws.append(['India','Asia'])          # add another row\n",
    "ws.append(['France','Europe'])       # add another row\n",
    "wb.save(\"countries.xlsx\")            # save the workbook into a file called countries.xlsx\n",
    "\n",
    "\n",
    "-->\n"
   ]
  },
  {
   "cell_type": "markdown",
   "metadata": {},
   "source": [
    "Create a python list of all locations for which you need to find the number of jobs postings.\n"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 19,
   "metadata": {
    "tags": []
   },
   "outputs": [
    {
     "name": "stdout",
     "output_type": "stream",
     "text": [
      "Collecting openpyxl\n",
      "  Downloading openpyxl-3.1.2-py2.py3-none-any.whl (249 kB)\n",
      "\u001b[2K     \u001b[90m━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\u001b[0m \u001b[32m250.0/250.0 kB\u001b[0m \u001b[31m29.4 MB/s\u001b[0m eta \u001b[36m0:00:00\u001b[0m\n",
      "\u001b[?25hCollecting et-xmlfile (from openpyxl)\n",
      "  Downloading et_xmlfile-1.1.0-py3-none-any.whl (4.7 kB)\n",
      "Installing collected packages: et-xmlfile, openpyxl\n",
      "Successfully installed et-xmlfile-1.1.0 openpyxl-3.1.2\n",
      "Note: you may need to restart the kernel to use updated packages.\n"
     ]
    }
   ],
   "source": [
    "pip install openpyxl\n"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 20,
   "metadata": {
    "tags": []
   },
   "outputs": [
    {
     "ename": "ConnectionError",
     "evalue": "HTTPSConnectionPool(host='api.example.com', port=443): Max retries exceeded with url: /jobs?technology=Python&location=New%20York (Caused by NewConnectionError('<urllib3.connection.HTTPSConnection object at 0x7f666dc57e90>: Failed to establish a new connection: [Errno -5] No address associated with hostname'))",
     "output_type": "error",
     "traceback": [
      "\u001b[0;31m---------------------------------------------------------------------------\u001b[0m",
      "\u001b[0;31mgaierror\u001b[0m                                  Traceback (most recent call last)",
      "\u001b[0;32m~/conda/envs/python/lib/python3.7/site-packages/urllib3/connection.py\u001b[0m in \u001b[0;36m_new_conn\u001b[0;34m(self)\u001b[0m\n\u001b[1;32m    174\u001b[0m             conn = connection.create_connection(\n\u001b[0;32m--> 175\u001b[0;31m                 \u001b[0;34m(\u001b[0m\u001b[0mself\u001b[0m\u001b[0;34m.\u001b[0m\u001b[0m_dns_host\u001b[0m\u001b[0;34m,\u001b[0m \u001b[0mself\u001b[0m\u001b[0;34m.\u001b[0m\u001b[0mport\u001b[0m\u001b[0;34m)\u001b[0m\u001b[0;34m,\u001b[0m \u001b[0mself\u001b[0m\u001b[0;34m.\u001b[0m\u001b[0mtimeout\u001b[0m\u001b[0;34m,\u001b[0m \u001b[0;34m**\u001b[0m\u001b[0mextra_kw\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0m\n\u001b[0m\u001b[1;32m    176\u001b[0m             )\n",
      "\u001b[0;32m~/conda/envs/python/lib/python3.7/site-packages/urllib3/util/connection.py\u001b[0m in \u001b[0;36mcreate_connection\u001b[0;34m(address, timeout, source_address, socket_options)\u001b[0m\n\u001b[1;32m     71\u001b[0m \u001b[0;34m\u001b[0m\u001b[0m\n\u001b[0;32m---> 72\u001b[0;31m     \u001b[0;32mfor\u001b[0m \u001b[0mres\u001b[0m \u001b[0;32min\u001b[0m \u001b[0msocket\u001b[0m\u001b[0;34m.\u001b[0m\u001b[0mgetaddrinfo\u001b[0m\u001b[0;34m(\u001b[0m\u001b[0mhost\u001b[0m\u001b[0;34m,\u001b[0m \u001b[0mport\u001b[0m\u001b[0;34m,\u001b[0m \u001b[0mfamily\u001b[0m\u001b[0;34m,\u001b[0m \u001b[0msocket\u001b[0m\u001b[0;34m.\u001b[0m\u001b[0mSOCK_STREAM\u001b[0m\u001b[0;34m)\u001b[0m\u001b[0;34m:\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0m\n\u001b[0m\u001b[1;32m     73\u001b[0m         \u001b[0maf\u001b[0m\u001b[0;34m,\u001b[0m \u001b[0msocktype\u001b[0m\u001b[0;34m,\u001b[0m \u001b[0mproto\u001b[0m\u001b[0;34m,\u001b[0m \u001b[0mcanonname\u001b[0m\u001b[0;34m,\u001b[0m \u001b[0msa\u001b[0m \u001b[0;34m=\u001b[0m \u001b[0mres\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0m\n",
      "\u001b[0;32m~/conda/envs/python/lib/python3.7/socket.py\u001b[0m in \u001b[0;36mgetaddrinfo\u001b[0;34m(host, port, family, type, proto, flags)\u001b[0m\n\u001b[1;32m    751\u001b[0m     \u001b[0maddrlist\u001b[0m \u001b[0;34m=\u001b[0m \u001b[0;34m[\u001b[0m\u001b[0;34m]\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0m\n\u001b[0;32m--> 752\u001b[0;31m     \u001b[0;32mfor\u001b[0m \u001b[0mres\u001b[0m \u001b[0;32min\u001b[0m \u001b[0m_socket\u001b[0m\u001b[0;34m.\u001b[0m\u001b[0mgetaddrinfo\u001b[0m\u001b[0;34m(\u001b[0m\u001b[0mhost\u001b[0m\u001b[0;34m,\u001b[0m \u001b[0mport\u001b[0m\u001b[0;34m,\u001b[0m \u001b[0mfamily\u001b[0m\u001b[0;34m,\u001b[0m \u001b[0mtype\u001b[0m\u001b[0;34m,\u001b[0m \u001b[0mproto\u001b[0m\u001b[0;34m,\u001b[0m \u001b[0mflags\u001b[0m\u001b[0;34m)\u001b[0m\u001b[0;34m:\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0m\n\u001b[0m\u001b[1;32m    753\u001b[0m         \u001b[0maf\u001b[0m\u001b[0;34m,\u001b[0m \u001b[0msocktype\u001b[0m\u001b[0;34m,\u001b[0m \u001b[0mproto\u001b[0m\u001b[0;34m,\u001b[0m \u001b[0mcanonname\u001b[0m\u001b[0;34m,\u001b[0m \u001b[0msa\u001b[0m \u001b[0;34m=\u001b[0m \u001b[0mres\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0m\n",
      "\u001b[0;31mgaierror\u001b[0m: [Errno -5] No address associated with hostname",
      "\nDuring handling of the above exception, another exception occurred:\n",
      "\u001b[0;31mNewConnectionError\u001b[0m                        Traceback (most recent call last)",
      "\u001b[0;32m~/conda/envs/python/lib/python3.7/site-packages/urllib3/connectionpool.py\u001b[0m in \u001b[0;36murlopen\u001b[0;34m(self, method, url, body, headers, retries, redirect, assert_same_host, timeout, pool_timeout, release_conn, chunked, body_pos, **response_kw)\u001b[0m\n\u001b[1;32m    709\u001b[0m                 \u001b[0mheaders\u001b[0m\u001b[0;34m=\u001b[0m\u001b[0mheaders\u001b[0m\u001b[0;34m,\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0m\n\u001b[0;32m--> 710\u001b[0;31m                 \u001b[0mchunked\u001b[0m\u001b[0;34m=\u001b[0m\u001b[0mchunked\u001b[0m\u001b[0;34m,\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0m\n\u001b[0m\u001b[1;32m    711\u001b[0m             )\n",
      "\u001b[0;32m~/conda/envs/python/lib/python3.7/site-packages/urllib3/connectionpool.py\u001b[0m in \u001b[0;36m_make_request\u001b[0;34m(self, conn, method, url, timeout, chunked, **httplib_request_kw)\u001b[0m\n\u001b[1;32m    385\u001b[0m         \u001b[0;32mtry\u001b[0m\u001b[0;34m:\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0m\n\u001b[0;32m--> 386\u001b[0;31m             \u001b[0mself\u001b[0m\u001b[0;34m.\u001b[0m\u001b[0m_validate_conn\u001b[0m\u001b[0;34m(\u001b[0m\u001b[0mconn\u001b[0m\u001b[0;34m)\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0m\n\u001b[0m\u001b[1;32m    387\u001b[0m         \u001b[0;32mexcept\u001b[0m \u001b[0;34m(\u001b[0m\u001b[0mSocketTimeout\u001b[0m\u001b[0;34m,\u001b[0m \u001b[0mBaseSSLError\u001b[0m\u001b[0;34m)\u001b[0m \u001b[0;32mas\u001b[0m \u001b[0me\u001b[0m\u001b[0;34m:\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0m\n",
      "\u001b[0;32m~/conda/envs/python/lib/python3.7/site-packages/urllib3/connectionpool.py\u001b[0m in \u001b[0;36m_validate_conn\u001b[0;34m(self, conn)\u001b[0m\n\u001b[1;32m   1041\u001b[0m         \u001b[0;32mif\u001b[0m \u001b[0;32mnot\u001b[0m \u001b[0mgetattr\u001b[0m\u001b[0;34m(\u001b[0m\u001b[0mconn\u001b[0m\u001b[0;34m,\u001b[0m \u001b[0;34m\"sock\"\u001b[0m\u001b[0;34m,\u001b[0m \u001b[0;32mNone\u001b[0m\u001b[0;34m)\u001b[0m\u001b[0;34m:\u001b[0m  \u001b[0;31m# AppEngine might not have  `.sock`\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0m\n\u001b[0;32m-> 1042\u001b[0;31m             \u001b[0mconn\u001b[0m\u001b[0;34m.\u001b[0m\u001b[0mconnect\u001b[0m\u001b[0;34m(\u001b[0m\u001b[0;34m)\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0m\n\u001b[0m\u001b[1;32m   1043\u001b[0m \u001b[0;34m\u001b[0m\u001b[0m\n",
      "\u001b[0;32m~/conda/envs/python/lib/python3.7/site-packages/urllib3/connection.py\u001b[0m in \u001b[0;36mconnect\u001b[0;34m(self)\u001b[0m\n\u001b[1;32m    362\u001b[0m         \u001b[0;31m# Add certificate verification\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0m\n\u001b[0;32m--> 363\u001b[0;31m         \u001b[0mself\u001b[0m\u001b[0;34m.\u001b[0m\u001b[0msock\u001b[0m \u001b[0;34m=\u001b[0m \u001b[0mconn\u001b[0m \u001b[0;34m=\u001b[0m \u001b[0mself\u001b[0m\u001b[0;34m.\u001b[0m\u001b[0m_new_conn\u001b[0m\u001b[0;34m(\u001b[0m\u001b[0;34m)\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0m\n\u001b[0m\u001b[1;32m    364\u001b[0m         \u001b[0mhostname\u001b[0m \u001b[0;34m=\u001b[0m \u001b[0mself\u001b[0m\u001b[0;34m.\u001b[0m\u001b[0mhost\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0m\n",
      "\u001b[0;32m~/conda/envs/python/lib/python3.7/site-packages/urllib3/connection.py\u001b[0m in \u001b[0;36m_new_conn\u001b[0;34m(self)\u001b[0m\n\u001b[1;32m    186\u001b[0m             raise NewConnectionError(\n\u001b[0;32m--> 187\u001b[0;31m                 \u001b[0mself\u001b[0m\u001b[0;34m,\u001b[0m \u001b[0;34m\"Failed to establish a new connection: %s\"\u001b[0m \u001b[0;34m%\u001b[0m \u001b[0me\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0m\n\u001b[0m\u001b[1;32m    188\u001b[0m             )\n",
      "\u001b[0;31mNewConnectionError\u001b[0m: <urllib3.connection.HTTPSConnection object at 0x7f666dc57e90>: Failed to establish a new connection: [Errno -5] No address associated with hostname",
      "\nDuring handling of the above exception, another exception occurred:\n",
      "\u001b[0;31mMaxRetryError\u001b[0m                             Traceback (most recent call last)",
      "\u001b[0;32m~/conda/envs/python/lib/python3.7/site-packages/requests/adapters.py\u001b[0m in \u001b[0;36msend\u001b[0;34m(self, request, stream, timeout, verify, cert, proxies)\u001b[0m\n\u001b[1;32m    497\u001b[0m                 \u001b[0mtimeout\u001b[0m\u001b[0;34m=\u001b[0m\u001b[0mtimeout\u001b[0m\u001b[0;34m,\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0m\n\u001b[0;32m--> 498\u001b[0;31m                 \u001b[0mchunked\u001b[0m\u001b[0;34m=\u001b[0m\u001b[0mchunked\u001b[0m\u001b[0;34m,\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0m\n\u001b[0m\u001b[1;32m    499\u001b[0m             )\n",
      "\u001b[0;32m~/conda/envs/python/lib/python3.7/site-packages/urllib3/connectionpool.py\u001b[0m in \u001b[0;36murlopen\u001b[0;34m(self, method, url, body, headers, retries, redirect, assert_same_host, timeout, pool_timeout, release_conn, chunked, body_pos, **response_kw)\u001b[0m\n\u001b[1;32m    787\u001b[0m             retries = retries.increment(\n\u001b[0;32m--> 788\u001b[0;31m                 \u001b[0mmethod\u001b[0m\u001b[0;34m,\u001b[0m \u001b[0murl\u001b[0m\u001b[0;34m,\u001b[0m \u001b[0merror\u001b[0m\u001b[0;34m=\u001b[0m\u001b[0me\u001b[0m\u001b[0;34m,\u001b[0m \u001b[0m_pool\u001b[0m\u001b[0;34m=\u001b[0m\u001b[0mself\u001b[0m\u001b[0;34m,\u001b[0m \u001b[0m_stacktrace\u001b[0m\u001b[0;34m=\u001b[0m\u001b[0msys\u001b[0m\u001b[0;34m.\u001b[0m\u001b[0mexc_info\u001b[0m\u001b[0;34m(\u001b[0m\u001b[0;34m)\u001b[0m\u001b[0;34m[\u001b[0m\u001b[0;36m2\u001b[0m\u001b[0;34m]\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0m\n\u001b[0m\u001b[1;32m    789\u001b[0m             )\n",
      "\u001b[0;32m~/conda/envs/python/lib/python3.7/site-packages/urllib3/util/retry.py\u001b[0m in \u001b[0;36mincrement\u001b[0;34m(self, method, url, response, error, _pool, _stacktrace)\u001b[0m\n\u001b[1;32m    591\u001b[0m         \u001b[0;32mif\u001b[0m \u001b[0mnew_retry\u001b[0m\u001b[0;34m.\u001b[0m\u001b[0mis_exhausted\u001b[0m\u001b[0;34m(\u001b[0m\u001b[0;34m)\u001b[0m\u001b[0;34m:\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0m\n\u001b[0;32m--> 592\u001b[0;31m             \u001b[0;32mraise\u001b[0m \u001b[0mMaxRetryError\u001b[0m\u001b[0;34m(\u001b[0m\u001b[0m_pool\u001b[0m\u001b[0;34m,\u001b[0m \u001b[0murl\u001b[0m\u001b[0;34m,\u001b[0m \u001b[0merror\u001b[0m \u001b[0;32mor\u001b[0m \u001b[0mResponseError\u001b[0m\u001b[0;34m(\u001b[0m\u001b[0mcause\u001b[0m\u001b[0;34m)\u001b[0m\u001b[0;34m)\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0m\n\u001b[0m\u001b[1;32m    593\u001b[0m \u001b[0;34m\u001b[0m\u001b[0m\n",
      "\u001b[0;31mMaxRetryError\u001b[0m: HTTPSConnectionPool(host='api.example.com', port=443): Max retries exceeded with url: /jobs?technology=Python&location=New%20York (Caused by NewConnectionError('<urllib3.connection.HTTPSConnection object at 0x7f666dc57e90>: Failed to establish a new connection: [Errno -5] No address associated with hostname'))",
      "\nDuring handling of the above exception, another exception occurred:\n",
      "\u001b[0;31mConnectionError\u001b[0m                           Traceback (most recent call last)",
      "\u001b[0;32m/tmp/ipykernel_911/2200865628.py\u001b[0m in \u001b[0;36m<module>\u001b[0;34m\u001b[0m\n\u001b[1;32m     28\u001b[0m     \u001b[0;32mfor\u001b[0m \u001b[0mlocation\u001b[0m \u001b[0;32min\u001b[0m \u001b[0mlocations\u001b[0m\u001b[0;34m:\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0m\n\u001b[1;32m     29\u001b[0m         \u001b[0;31m# Call the API\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0m\n\u001b[0;32m---> 30\u001b[0;31m         \u001b[0mdata\u001b[0m \u001b[0;34m=\u001b[0m \u001b[0mcall_api\u001b[0m\u001b[0;34m(\u001b[0m\u001b[0mtechnology\u001b[0m\u001b[0;34m,\u001b[0m \u001b[0mlocation\u001b[0m\u001b[0;34m)\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0m\n\u001b[0m\u001b[1;32m     31\u001b[0m         \u001b[0;32mif\u001b[0m \u001b[0mdata\u001b[0m\u001b[0;34m:\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0m\n\u001b[1;32m     32\u001b[0m             \u001b[0;31m# Extract the number of job postings\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0m\n",
      "\u001b[0;32m/tmp/ipykernel_911/2200865628.py\u001b[0m in \u001b[0;36mcall_api\u001b[0;34m(technology, location)\u001b[0m\n\u001b[1;32m     11\u001b[0m \u001b[0;32mdef\u001b[0m \u001b[0mcall_api\u001b[0m\u001b[0;34m(\u001b[0m\u001b[0mtechnology\u001b[0m\u001b[0;34m,\u001b[0m \u001b[0mlocation\u001b[0m\u001b[0;34m)\u001b[0m\u001b[0;34m:\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0m\n\u001b[1;32m     12\u001b[0m     \u001b[0murl\u001b[0m \u001b[0;34m=\u001b[0m \u001b[0;34mf\"https://api.example.com/jobs?technology={technology}&location={location}\"\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0m\n\u001b[0;32m---> 13\u001b[0;31m     \u001b[0mresponse\u001b[0m \u001b[0;34m=\u001b[0m \u001b[0mrequests\u001b[0m\u001b[0;34m.\u001b[0m\u001b[0mget\u001b[0m\u001b[0;34m(\u001b[0m\u001b[0murl\u001b[0m\u001b[0;34m)\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0m\n\u001b[0m\u001b[1;32m     14\u001b[0m     \u001b[0;32mif\u001b[0m \u001b[0mresponse\u001b[0m\u001b[0;34m.\u001b[0m\u001b[0mstatus_code\u001b[0m \u001b[0;34m==\u001b[0m \u001b[0;36m200\u001b[0m\u001b[0;34m:\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0m\n\u001b[1;32m     15\u001b[0m         \u001b[0;32mreturn\u001b[0m \u001b[0mresponse\u001b[0m\u001b[0;34m.\u001b[0m\u001b[0mjson\u001b[0m\u001b[0;34m(\u001b[0m\u001b[0;34m)\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0m\n",
      "\u001b[0;32m~/conda/envs/python/lib/python3.7/site-packages/requests/api.py\u001b[0m in \u001b[0;36mget\u001b[0;34m(url, params, **kwargs)\u001b[0m\n\u001b[1;32m     71\u001b[0m     \"\"\"\n\u001b[1;32m     72\u001b[0m \u001b[0;34m\u001b[0m\u001b[0m\n\u001b[0;32m---> 73\u001b[0;31m     \u001b[0;32mreturn\u001b[0m \u001b[0mrequest\u001b[0m\u001b[0;34m(\u001b[0m\u001b[0;34m\"get\"\u001b[0m\u001b[0;34m,\u001b[0m \u001b[0murl\u001b[0m\u001b[0;34m,\u001b[0m \u001b[0mparams\u001b[0m\u001b[0;34m=\u001b[0m\u001b[0mparams\u001b[0m\u001b[0;34m,\u001b[0m \u001b[0;34m**\u001b[0m\u001b[0mkwargs\u001b[0m\u001b[0;34m)\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0m\n\u001b[0m\u001b[1;32m     74\u001b[0m \u001b[0;34m\u001b[0m\u001b[0m\n\u001b[1;32m     75\u001b[0m \u001b[0;34m\u001b[0m\u001b[0m\n",
      "\u001b[0;32m~/conda/envs/python/lib/python3.7/site-packages/requests/api.py\u001b[0m in \u001b[0;36mrequest\u001b[0;34m(method, url, **kwargs)\u001b[0m\n\u001b[1;32m     57\u001b[0m     \u001b[0;31m# cases, and look like a memory leak in others.\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0m\n\u001b[1;32m     58\u001b[0m     \u001b[0;32mwith\u001b[0m \u001b[0msessions\u001b[0m\u001b[0;34m.\u001b[0m\u001b[0mSession\u001b[0m\u001b[0;34m(\u001b[0m\u001b[0;34m)\u001b[0m \u001b[0;32mas\u001b[0m \u001b[0msession\u001b[0m\u001b[0;34m:\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0m\n\u001b[0;32m---> 59\u001b[0;31m         \u001b[0;32mreturn\u001b[0m \u001b[0msession\u001b[0m\u001b[0;34m.\u001b[0m\u001b[0mrequest\u001b[0m\u001b[0;34m(\u001b[0m\u001b[0mmethod\u001b[0m\u001b[0;34m=\u001b[0m\u001b[0mmethod\u001b[0m\u001b[0;34m,\u001b[0m \u001b[0murl\u001b[0m\u001b[0;34m=\u001b[0m\u001b[0murl\u001b[0m\u001b[0;34m,\u001b[0m \u001b[0;34m**\u001b[0m\u001b[0mkwargs\u001b[0m\u001b[0;34m)\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0m\n\u001b[0m\u001b[1;32m     60\u001b[0m \u001b[0;34m\u001b[0m\u001b[0m\n\u001b[1;32m     61\u001b[0m \u001b[0;34m\u001b[0m\u001b[0m\n",
      "\u001b[0;32m~/conda/envs/python/lib/python3.7/site-packages/requests/sessions.py\u001b[0m in \u001b[0;36mrequest\u001b[0;34m(self, method, url, params, data, headers, cookies, files, auth, timeout, allow_redirects, proxies, hooks, stream, verify, cert, json)\u001b[0m\n\u001b[1;32m    585\u001b[0m         }\n\u001b[1;32m    586\u001b[0m         \u001b[0msend_kwargs\u001b[0m\u001b[0;34m.\u001b[0m\u001b[0mupdate\u001b[0m\u001b[0;34m(\u001b[0m\u001b[0msettings\u001b[0m\u001b[0;34m)\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0m\n\u001b[0;32m--> 587\u001b[0;31m         \u001b[0mresp\u001b[0m \u001b[0;34m=\u001b[0m \u001b[0mself\u001b[0m\u001b[0;34m.\u001b[0m\u001b[0msend\u001b[0m\u001b[0;34m(\u001b[0m\u001b[0mprep\u001b[0m\u001b[0;34m,\u001b[0m \u001b[0;34m**\u001b[0m\u001b[0msend_kwargs\u001b[0m\u001b[0;34m)\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0m\n\u001b[0m\u001b[1;32m    588\u001b[0m \u001b[0;34m\u001b[0m\u001b[0m\n\u001b[1;32m    589\u001b[0m         \u001b[0;32mreturn\u001b[0m \u001b[0mresp\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0m\n",
      "\u001b[0;32m~/conda/envs/python/lib/python3.7/site-packages/requests/sessions.py\u001b[0m in \u001b[0;36msend\u001b[0;34m(self, request, **kwargs)\u001b[0m\n\u001b[1;32m    699\u001b[0m \u001b[0;34m\u001b[0m\u001b[0m\n\u001b[1;32m    700\u001b[0m         \u001b[0;31m# Send the request\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0m\n\u001b[0;32m--> 701\u001b[0;31m         \u001b[0mr\u001b[0m \u001b[0;34m=\u001b[0m \u001b[0madapter\u001b[0m\u001b[0;34m.\u001b[0m\u001b[0msend\u001b[0m\u001b[0;34m(\u001b[0m\u001b[0mrequest\u001b[0m\u001b[0;34m,\u001b[0m \u001b[0;34m**\u001b[0m\u001b[0mkwargs\u001b[0m\u001b[0;34m)\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0m\n\u001b[0m\u001b[1;32m    702\u001b[0m \u001b[0;34m\u001b[0m\u001b[0m\n\u001b[1;32m    703\u001b[0m         \u001b[0;31m# Total elapsed time of the request (approximately)\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0m\n",
      "\u001b[0;32m~/conda/envs/python/lib/python3.7/site-packages/requests/adapters.py\u001b[0m in \u001b[0;36msend\u001b[0;34m(self, request, stream, timeout, verify, cert, proxies)\u001b[0m\n\u001b[1;32m    518\u001b[0m                 \u001b[0;32mraise\u001b[0m \u001b[0mSSLError\u001b[0m\u001b[0;34m(\u001b[0m\u001b[0me\u001b[0m\u001b[0;34m,\u001b[0m \u001b[0mrequest\u001b[0m\u001b[0;34m=\u001b[0m\u001b[0mrequest\u001b[0m\u001b[0;34m)\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0m\n\u001b[1;32m    519\u001b[0m \u001b[0;34m\u001b[0m\u001b[0m\n\u001b[0;32m--> 520\u001b[0;31m             \u001b[0;32mraise\u001b[0m \u001b[0mConnectionError\u001b[0m\u001b[0;34m(\u001b[0m\u001b[0me\u001b[0m\u001b[0;34m,\u001b[0m \u001b[0mrequest\u001b[0m\u001b[0;34m=\u001b[0m\u001b[0mrequest\u001b[0m\u001b[0;34m)\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0m\n\u001b[0m\u001b[1;32m    521\u001b[0m \u001b[0;34m\u001b[0m\u001b[0m\n\u001b[1;32m    522\u001b[0m         \u001b[0;32mexcept\u001b[0m \u001b[0mClosedPoolError\u001b[0m \u001b[0;32mas\u001b[0m \u001b[0me\u001b[0m\u001b[0;34m:\u001b[0m\u001b[0;34m\u001b[0m\u001b[0;34m\u001b[0m\u001b[0m\n",
      "\u001b[0;31mConnectionError\u001b[0m: HTTPSConnectionPool(host='api.example.com', port=443): Max retries exceeded with url: /jobs?technology=Python&location=New%20York (Caused by NewConnectionError('<urllib3.connection.HTTPSConnection object at 0x7f666dc57e90>: Failed to establish a new connection: [Errno -5] No address associated with hostname'))"
     ]
    }
   ],
   "source": [
    "import requests\n",
    "from openpyxl import Workbook\n",
    "\n",
    "# List of technologies\n",
    "technologies = [\"Python\", \"Java\", \"JavaScript\", \"Ruby\"]\n",
    "\n",
    "# List of locations\n",
    "locations = [\"New York\", \"San Francisco\", \"London\", \"Tokyo\"]\n",
    "\n",
    "# Function to call the API for a given technology and location\n",
    "def call_api(technology, location):\n",
    "    url = f\"https://api.example.com/jobs?technology={technology}&location={location}\"\n",
    "    response = requests.get(url)\n",
    "    if response.status_code == 200:\n",
    "        return response.json()\n",
    "    else:\n",
    "        return None\n",
    "\n",
    "# Create a new Excel workbook\n",
    "wb = Workbook()\n",
    "ws = wb.active\n",
    "\n",
    "# Write headers to the first row\n",
    "ws.append([\"Technology\", \"Location\", \"Job Postings\"])\n",
    "\n",
    "# Iterate over each technology and location\n",
    "for technology in technologies:\n",
    "    for location in locations:\n",
    "        # Call the API\n",
    "        data = call_api(technology, location)\n",
    "        if data:\n",
    "            # Extract the number of job postings\n",
    "            job_postings = data.get(\"total_jobs\")\n",
    "            # Write data to Excel\n",
    "            ws.append([technology, location, job_postings])\n",
    "\n",
    "# Save the workbook\n",
    "wb.save(\"job_postings.xlsx\")\n"
   ]
  },
  {
   "cell_type": "markdown",
   "metadata": {},
   "source": [
    "Import libraries required to create excel spreadsheet\n"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 14,
   "metadata": {
    "tags": []
   },
   "outputs": [],
   "source": [
    "# your code goes here"
   ]
  },
  {
   "cell_type": "markdown",
   "metadata": {},
   "source": [
    "Create a workbook and select the active worksheet\n"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": null,
   "metadata": {},
   "outputs": [],
   "source": [
    "# your code goes here"
   ]
  },
  {
   "cell_type": "markdown",
   "metadata": {},
   "source": [
    "Find the number of jobs postings for each of the location in the above list.\n",
    "Write the Location name and the number of jobs postings into the excel spreadsheet.\n"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": null,
   "metadata": {},
   "outputs": [],
   "source": [
    "#your code goes here"
   ]
  },
  {
   "cell_type": "markdown",
   "metadata": {},
   "source": [
    "Save into an excel spreadsheet named 'job-postings.xlsx'.\n"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": null,
   "metadata": {},
   "outputs": [],
   "source": [
    "#your code goes here"
   ]
  },
  {
   "cell_type": "markdown",
   "metadata": {},
   "source": [
    "#### In the similar way, you can try for below given technologies and results  can be stored in an excel sheet.\n"
   ]
  },
  {
   "cell_type": "markdown",
   "metadata": {},
   "source": [
    "Collect the number of job postings for the following languages using the API:\n",
    "\n",
    "*   C\n",
    "*   C#\n",
    "*   C++\n",
    "*   Java\n",
    "*   JavaScript\n",
    "*   Python\n",
    "*   Scala\n",
    "*   Oracle\n",
    "*   SQL Server\n",
    "*   MySQL Server\n",
    "*   PostgreSQL\n",
    "*   MongoDB\n"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": null,
   "metadata": {},
   "outputs": [],
   "source": [
    "# your code goes here\n"
   ]
  },
  {
   "cell_type": "markdown",
   "metadata": {},
   "source": [
    "## Author\n"
   ]
  },
  {
   "cell_type": "markdown",
   "metadata": {},
   "source": [
    "Ayushi Jain\n"
   ]
  },
  {
   "cell_type": "markdown",
   "metadata": {},
   "source": [
    "### Other Contributors\n"
   ]
  },
  {
   "cell_type": "markdown",
   "metadata": {},
   "source": [
    "Rav Ahuja\n",
    "\n",
    "Lakshmi Holla\n",
    "\n",
    "Malika\n"
   ]
  },
  {
   "cell_type": "markdown",
   "metadata": {},
   "source": [
    "## Change Log\n"
   ]
  },
  {
   "cell_type": "markdown",
   "metadata": {},
   "source": [
    "| Date (YYYY-MM-DD) | Version | Changed By        | Change Description                 |\n",
    "| ----------------- | ------- | ----------------- | ---------------------------------- | \n",
    "| 2022-01-19        | 0.3     | Lakshmi Holla        | Added changes in the markdown      |\n",
    "| 2021-06-25        | 0.2     | Malika            | Updated GitHub job json link       |\n",
    "| 2020-10-17        | 0.1     | Ramesh Sannareddy | Created initial version of the lab |\n"
   ]
  },
  {
   "cell_type": "markdown",
   "metadata": {},
   "source": [
    "Copyright © 2022 IBM Corporation. All rights reserved. \n"
   ]
  }
 ],
 "metadata": {
  "kernelspec": {
   "display_name": "Python",
   "language": "python",
   "name": "conda-env-python-py"
  },
  "language_info": {
   "codemirror_mode": {
    "name": "ipython",
    "version": 3
   },
   "file_extension": ".py",
   "mimetype": "text/x-python",
   "name": "python",
   "nbconvert_exporter": "python",
   "pygments_lexer": "ipython3",
   "version": "3.7.12"
  }
 },
 "nbformat": 4,
 "nbformat_minor": 4
}
