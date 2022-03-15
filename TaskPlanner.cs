using System;
using System.Linq;
using System.Windows.Forms;

namespace TaskPlanner_Build_2
{
    using System.Drawing;
    using System.IO;
    using System.Threading;

    public partial class TaskPlanner : Form
    {
        public TaskPlanner()
        {
            InitializeComponent();
            this.TopMost = false;
            //Populate Project list box with the names of the projects in Project Listings.txt upon startup
            fileLocation = Application.StartupPath;
            string[] projects = File.ReadAllLines($@"{fileLocation}\Projects\Project Listings.txt");
            foreach (string project in projects)
                if (project != "")
                    projectListBox.Items.Add(project);
        }

        //Task text file located here:  $"{project.fileLocation}\\Projects\\{project.projectListBox.SelectedItem.ToString()}\\{project.taskListBox.SelectedItem.ToString()}\\{project.taskListBox.SelectedItem.ToString()}.txt"
        //or here:  $@"{fileLocation}\Projects\{projectListBox.SelectedItem.ToString()}\{taskListBox.SelectedItem.ToString()}\{taskListBox.SelectedItem.ToString()}.txt"

        //Holds save file location
        public string fileLocation;

        public void TaskDataClearer()
        {
            //Clears all task-related data on the form
            membersListBox.Items.Clear();
            messageListBox.Items.Clear();
            stepsCheckedListBox.Items.Clear();
            descriptionRichTextBox.Text = "";
            taskNameLabel.Text = "Task Name:";
            startTextBox.Text = "";
            endTextBox.Text = "";
            //Reset status button colors
            normalButton.BackColor = SystemColors.Control;
            stuckButton.BackColor = SystemColors.Control;
            urgentButton.BackColor = SystemColors.Control;
            completeButton.BackColor = SystemColors.Control;
        }

        //Holds the task name
        public string taskName;
        public void TaskNameFinder()
        {
            //Finds the name of the selected task from the project text file
            string fileName = $@"{fileLocation}\Projects\{projectListBox.SelectedItem.ToString()}\{projectListBox.SelectedItem.ToString()}.txt";
            string tempString = File.ReadLines(fileName).Skip(taskListBox.SelectedIndex).Take(1).First();
            string[] taskNameArray = tempString.Split('|');
            taskName = taskNameArray[0];
        }

        public void LineChanger(string newText, int lineToEdit)
        {
            //Writes and overwrites data to specified lines in the text files.
            TaskNameFinder();
            string fileName = $@"{fileLocation}\Projects\{projectListBox.SelectedItem.ToString()}\{taskName}\{taskName}.txt";
            string[] allLines = File.ReadAllLines(fileName);
            allLines[lineToEdit - 1] = newText;
            File.WriteAllLines(fileName, allLines);
        }

        public void LineChangerNoOverwrite(string newText, int lineToEdit)
        {
            //Appends data to specific lines in the text files without overwriting.
            TaskNameFinder();
            string fileName = $@"{fileLocation}\Projects\{projectListBox.SelectedItem.ToString()}\{taskName.ToString()}\{taskName.ToString()}.txt";
            string[] allLines = File.ReadAllLines(fileName);
            //Checks line to be overwritten to see if it contains anything.  If it does, it formats the line using '|' (pipe) as a delimiter
            string checkLine = File.ReadLines(fileName).Skip(lineToEdit - 1).Take(1).First();
            if (checkLine != "")
            {
                allLines[lineToEdit - 1] += "|" + newText;
                File.WriteAllLines(fileName, allLines);
            }
            else
            {
                allLines[lineToEdit - 1] = newText;
                File.WriteAllLines(fileName, allLines);
            }
        }



        //Pop up dialogue boxes to enter names/data=======================================================================================
        private void ProjectCreateButton_Click(object sender, EventArgs e)
        {
            //To create a new project
            ProjectPopUp PopUp = new ProjectPopUp();
            PopUp.ShowDialog();
        }

        private void TaskCreateButton_Click(object sender, EventArgs e)
        {
            //To create a new task
            if (projectListBox.SelectedIndex > -1)
            {
                TaskPopUp PopUp = new TaskPopUp();
                PopUp.ShowDialog();
            }
            else
                MessageBox.Show("Select a project to add the task to", "Select a Project", MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
        }

        private void AddMembersButton_Click(object sender, EventArgs e)
        {
            //To add a new member
            if (taskListBox.SelectedItem != null)
            {
                MemberPopUp PopUp = new MemberPopUp();
                PopUp.ShowDialog();
            }
            else
                MessageBox.Show("Select a task first", "", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        private void AddMessageButton_Click(object sender, EventArgs e)
        {
            //To create a new message
            if (taskListBox.SelectedItem != null)
            {
                MessageNamingPopUp PopUp = new MessageNamingPopUp();
                PopUp.ShowDialog();
            }
            else
                MessageBox.Show("Select a task first", "", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        private void MessageListBox_DoubleClick(object sender, EventArgs e)
        {
            //Opens the message-viewing form
            if (messageListBox.SelectedIndex > -1)
            {
                MessageContents EditView = new MessageContents();
                EditView.ShowDialog();
            }
        }

        private void AddStepButton_Click(object sender, EventArgs e)
        {
            //Add a new step
            if (taskListBox.SelectedItem != null)
            {
                StepsPopUp PopUp = new StepsPopUp();
                PopUp.ShowDialog();
            }
            else
                MessageBox.Show("Select a task first", "", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        private void DescriptionRichTextBox_DoubleClick(object sender, EventArgs e)
        {
            //Opens the task description-viewing form
            if (taskListBox.SelectedItem != null)
            {
                TaskDescriptionPopUp PopUp = new TaskDescriptionPopUp();
                PopUp.ShowDialog();
            }
            else
                MessageBox.Show("Select a task first", "", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        private void DateSortButton_Click(object sender, EventArgs e)
        {
            //Opens the task-sorter form
            TaskDateSorter PopUp = new TaskDateSorter();
            PopUp.ShowDialog();
        }

        private void MoveTaskButton_Click(object sender, EventArgs e)
        {
            //Move the selected task from one project to another
            if (taskListBox.SelectedIndex > -1)
            {
                TaskMoverPopUp PopUp = new TaskMoverPopUp();
                PopUp.ShowDialog();
            }
            else
                MessageBox.Show("Select a task first", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
        }

        //The next four methods reference DateClick() at the bottom to handle any date-related control being clicked
        private void StartDateLabel_Click(object sender, EventArgs e)
        {
            DateClick();
        }

        private void EndDateLabel_Click(object sender, EventArgs e)
        {
            DateClick();
        }

        private void StartTextBox_Click(object sender, EventArgs e)
        {
            DateClick();
        }

        private void EndTextBox_Click(object sender, EventArgs e)
        {
            DateClick();
        }

        public void DateClick()
        {
            //To change the start and end dates
            if (taskListBox.SelectedIndex > -1)
            {
                DatePopUp PopUp = new DatePopUp();
                PopUp.ShowDialog();
            }
            else
                MessageBox.Show("Select a task first.", "", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }




        //Add items from pop-up dialogues to the appropriate listboxes and update files and directories=========================================================================
        public void UpdateProjectListBox(string projectName)
        {
            //Sets up the project folders and files
            projectName = projectName.Trim();
            projectListBox.Items.Add(projectName);
            //Create project directory/file structure
            try
            {
                Directory.CreateDirectory($@"{fileLocation}\Projects\{projectName}");
                File.CreateText($@"{fileLocation}\Projects\{projectName}\{projectName}.txt").Dispose();
            }
            catch (IOException)
            {
                MessageBox.Show("Unable to create a new project.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            //Add project name to Project Listings.txt
            using (StreamWriter file = new StreamWriter($@"{fileLocation}\Projects\Project Listings.txt", true))
            {
                file.WriteLine(projectName);
            }
            //Alphabetize the Project Listings text file
            string[] projectAlphabetizer = File.ReadAllLines($@"{fileLocation}\Projects\Project Listings.txt");
            Array.Sort(projectAlphabetizer);
            File.WriteAllLines($@"{fileLocation}\Projects\Project Listings.txt", projectAlphabetizer);
        }

        public void UpdateTaskListBox(string taskName)
        {
            //Sets up the task folders and files
            taskName = taskName.Trim();
            taskListBox.Items.Add(taskName);
            //Create task directory/file structure
            try
            {
                Directory.CreateDirectory($@"{fileLocation}\Projects\{projectListBox.SelectedItem}\{taskName}");
                Directory.CreateDirectory($@"{fileLocation}\Projects\{projectListBox.SelectedItem}\{taskName}\Messages");
                File.CreateText($@"{fileLocation}\Projects\{projectListBox.SelectedItem}\{taskName}\{taskName}.txt").Dispose();
            }
            catch (IOException)
            {
                MessageBox.Show("Unable to create a new task.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            //Initialize lines in task text
            string taskTextPath = $@"{fileLocation}\Projects\{projectListBox.SelectedItem.ToString()}\{taskName}\{taskName}.txt";
            string[] textIni = { "//Editing these text files manually may make them unreadable", "|", "", "", "", "", "" };
            File.WriteAllLines(taskTextPath, textIni);
            //Add task name to project's task listings
            string projectTextPath = $@"{fileLocation}\Projects\{projectListBox.SelectedItem}\{projectListBox.SelectedItem}.txt";
            using (StreamWriter file = new StreamWriter(projectTextPath, true))
            {
                file.WriteLine(taskName + "|0");
            }
            //Alphabetize the tasks in the project's text file
            string[] taskAlphabetizer = File.ReadAllLines(projectTextPath);
            Array.Sort(taskAlphabetizer);
            File.WriteAllLines(projectTextPath, taskAlphabetizer);
        }

        public void UpdateMemberListBox(string memberName)
        {
            //Adds new member name to list box
            memberName = memberName.Trim();
            membersListBox.Items.Add(memberName);
        }

        public void UpdateMessageListBox(string messageName)
        {
            //Adds new message name to list box
            messageName = messageName.Trim();
            messageListBox.Items.Add(messageName);
        }

        public void UpdateStepsListBox(string stepName)
        {
            //Adds new step name to the checked list box
            stepName = stepName.Trim();
            stepsCheckedListBox.Items.Add(stepName);
        }

        public void UpdateDatesLabels(string datesLine)
        {
            //Displays the new dates
            string[] datesArray = datesLine.Split('|');
            startTextBox.Text = (datesArray[0]);
            endTextBox.Text = (datesArray[1]);
        }




        //Delete Project/Task/Member/Step/Message===========================================================================================
        private void ProjectDeleteButton_Click(object sender, EventArgs e)
        {
            //Deletes the selected project and all its tasks.
            if (projectListBox.SelectedIndex > -1)
            {
                DialogResult result = MessageBox.Show("Permanently delete this project and all its tasks? (No undo)", "Confirm Delete", MessageBoxButtons.YesNo,
                    MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2);

                if (result == DialogResult.Yes)
                {
                    //Delete project from Project Listings text file
                    string listingsFileLocation = $"{fileLocation}\\Projects\\Project Listings.txt";
                    string[] projects = File.ReadAllLines(listingsFileLocation);
                    using (StreamWriter file = new StreamWriter(listingsFileLocation))
                    {
                        for (int x = 0; x < projectListBox.Items.Count; ++x)
                        {
                            if (projects[x] != projectListBox.SelectedItem.ToString())
                                file.WriteLine(projects[x]);
                        }
                    }
                    //Delete whole project directory
                    try
                    {
                        Directory.Delete($@"{fileLocation}\Projects\{projectListBox.SelectedItem}", true);
                        projectListBox.Items.RemoveAt(projectListBox.SelectedIndex);
                        TaskDataClearer();
                        taskListBox.Items.Clear();
                    }
                    catch (IOException)
                    {
                        MessageBox.Show("Folder inaccessible.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
            }
            else
                MessageBox.Show("Select a project first", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
        }

        private void TaskDeleteButton_Click(object sender, EventArgs e)
        {
            //Deletes the selected task and all its messages
            if (taskListBox.SelectedIndex > -1)
            {
                DialogResult result = MessageBox.Show("Permanently delete this task? (No undo)", "Confirm Delete", MessageBoxButtons.YesNo,
                     MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2);
                if (result == DialogResult.Yes)
                {
                    //Delete task from the project text file
                    string projectFileLocation = $"{fileLocation}\\Projects\\{projectListBox.SelectedItem}\\{projectListBox.SelectedItem}.txt";
                    string[] tasks = File.ReadAllLines(projectFileLocation);
                    using (StreamWriter file = new StreamWriter(projectFileLocation))
                    {
                        for (int x = 0; x < taskListBox.Items.Count; ++x)
                        {
                            if (x != taskListBox.SelectedIndex)
                                file.WriteLine(tasks[x]);
                        }
                    }
                    //Delete whole task directory
                    try
                    {
                        Directory.Delete($@"{fileLocation}\Projects\{projectListBox.SelectedItem}\{taskName}", true);
                        taskListBox.Items.RemoveAt(taskListBox.SelectedIndex);
                        TaskDataClearer();
                    }
                    catch (IOException)
                    {
                        MessageBox.Show("Folder inaccessible.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
            }
            else
                MessageBox.Show("Select a task first", "Select a Task", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
        }

        private void RemoveMembersButton_Click(object sender, EventArgs e)
        {
            //Delete the selected member from the task text file
            if (membersListBox.SelectedIndex > -1)
            {
                DialogResult result = MessageBox.Show("Permanently delete this member? (No undo)", "Confirm Delete", MessageBoxButtons.YesNo,
                    MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2);
                if (result == DialogResult.Yes)
                {
                    //Delete member from the task text file
                    string taskTextLocation = $"{fileLocation}\\Projects\\{projectListBox.SelectedItem.ToString()}\\{taskName}\\{taskName}.txt";
                    string allMembers = File.ReadAllLines(taskTextLocation).Skip(2).Take(1).First();
                    string[] membersArray = allMembers.Split('|');
                    string newMembers = "";
                    for (int x = 0; x < membersArray.Count() - 1; ++x)
                    {
                        if (membersArray[x] != membersListBox.SelectedItem.ToString() && membersArray[x] != "")
                            newMembers += membersArray[x] + "|";
                    }
                    LineChanger(newMembers, 3);
                    membersListBox.Items.RemoveAt(membersListBox.SelectedIndex);
                }
            }
            else
                MessageBox.Show("Select a member", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
        }

        private void RemoveMessageButton_Click(object sender, EventArgs e)
        {
            //Deletes the selected message from the message list box and the task text file
            //Also deletes the selected message's text file from the selected task's folder
            if (messageListBox.SelectedIndex > -1)
            {
                DialogResult result = MessageBox.Show("Permanently delete this message and all its contents? (No undo)", "Confirm Delete", MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2);
                if (result == DialogResult.Yes)
                {
                    //Delete message from the task text file
                    string taskTextLocation = $"{fileLocation}\\Projects\\{projectListBox.SelectedItem.ToString()}\\{taskName}\\{taskName}.txt";
                    string allMessages = File.ReadAllLines(taskTextLocation).Skip(3).Take(1).First();
                    string[] messagesArray = allMessages.Split('|');
                    string newMessages = "";
                    for (int x = 0; x < messagesArray.Count(); ++x)
                    {
                        if (x != messageListBox.SelectedIndex && messagesArray[x] != "")
                            newMessages += messagesArray[x] + "|";
                    }
                    LineChanger(newMessages, 4);
                    //Delete message contents text file
                    try
                    {
                        File.Delete($@"{fileLocation}\Projects\{projectListBox.SelectedItem.ToString()}\{taskName}\Messages\{messageListBox.SelectedItem}.txt");
                        messageListBox.Items.RemoveAt(messageListBox.SelectedIndex);
                    }
                    catch (IOException)
                    {
                        MessageBox.Show("File inaccessible.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }
            }
            else
                MessageBox.Show("Select a message", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
        }

        private void RemoveStepButton_Click(object sender, EventArgs e)
        {
            //Deletes the selected step from the task text file
            //Also resets the list of checked steps in the task text file to reflect the new, shortened list of steps
            if (stepsCheckedListBox.SelectedIndex > -1)
            {
                DialogResult result = MessageBox.Show("Permanently delete this step? (No undo)", "Confirm Delete", MessageBoxButtons.YesNo,
                    MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2);
                if (result == DialogResult.Yes)
                {
                    //Delete step from the task text file
                    string taskTextLocation = $"{fileLocation}\\Projects\\{projectListBox.SelectedItem.ToString()}\\{taskName}\\{taskName}.txt";
                    string allSteps = File.ReadAllLines(taskTextLocation).Skip(4).Take(1).First();
                    string[] stepsArray = allSteps.Split('|');
                    string newSteps = "";
                    for (int x = 0; x < stepsArray.Count(); ++x)
                    {
                        if (x != stepsCheckedListBox.SelectedIndex && stepsArray[x] != "")
                            newSteps += stepsArray[x] + "|";
                    }
                    LineChanger(newSteps, 5);
                    stepsCheckedListBox.Items.RemoveAt(stepsCheckedListBox.SelectedIndex);
                    //Send new list of checked steps to text file
                    string checkedIndices = "";
                    foreach (int checkedIndexInt in stepsCheckedListBox.CheckedIndices)
                        checkedIndices += checkedIndexInt.ToString() + "|";
                    LineChanger(checkedIndices, 6);
                }
            }
            else
                MessageBox.Show("Select a step", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
        }




        //Populate data fields from text files===============================================================================
        private void ProjectListBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            //Populates the task list box with the selected project's tasks
            //Also appends a short string to the task name that tells the task's status
            if (projectListBox.SelectedItem != null)
            {
                //Clear all task data to reset the form
                taskListBox.Items.Clear();
                TaskDataClearer();
                string[] tasks = File.ReadAllLines($@"{fileLocation}\Projects\{projectListBox.SelectedItem}\{projectListBox.SelectedItem}.txt");
                string[] taskNameArr;
                string taskName;
                //Set task status in taskListBox
                foreach (string task in tasks)
                {
                    taskNameArr = task.Split('|');
                    switch (int.Parse(taskNameArr[1]))
                    {
                        case 0:
                            taskListBox.Items.Add(taskNameArr[0]);
                            break;
                        case 1:
                            taskName = taskNameArr[0] + "   ||   STUCK";
                            taskListBox.Items.Add(taskName);
                            break;
                        case 2:
                            taskName = taskNameArr[0] + "   ||   URGENT";
                            taskListBox.Items.Add(taskName);
                            break;
                        case 3:
                            taskName = taskNameArr[0] + "   ||   COMPLETE";
                            taskListBox.Items.Add(taskName);
                            break;
                        default:
                            taskListBox.Items.Add(taskNameArr[0]);
                            break;
                    }
                }
            }
        }

        private void TaskListBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            //Populates all task-related data on the form
            //This includes the task name, the start and end dates, the members, the messages, the steps, 
            //the status button color, and the task description.
            //Nearly all the information the user sees on the TaskPlanner form is loaded here
            if (taskListBox.SelectedItem != null)
            {
                //Determine the task name and file location
                TaskNameFinder();
                string taskTextLocation = $"{fileLocation}\\Projects\\{projectListBox.SelectedItem.ToString()}\\{taskName}\\{taskName}.txt";

                //Populate task name label
                taskNameLabel.Text = ($"Task Name:  {taskName}");

                //Populate dates text boxes
                string datesLine = File.ReadLines(taskTextLocation).Skip(1).First();
                if (datesLine != "|")
                {
                    string[] datesArr = datesLine.Split('|');
                    startTextBox.Text = datesArr[0];
                    endTextBox.Text = datesArr[1];
                }
                else
                {
                    startTextBox.Text = "(Click to Change)";
                    endTextBox.Text = "(Click to Change)";
                }

                //Populate members list box
                membersListBox.Items.Clear();
                string membersLine = File.ReadLines(taskTextLocation).Skip(2).Take(1).First();
                string[] membersArr = membersLine.Split('|');
                foreach (string membersString in membersArr)
                    if (membersString != "")
                        membersListBox.Items.Add(membersString);

                //Populate message list box
                messageListBox.Items.Clear();
                string messagesLine = File.ReadLines(taskTextLocation).Skip(3).Take(1).First();
                string[] messagesArr = messagesLine.Split('|');
                foreach (string messageString in messagesArr)
                    if (messageString != "")
                        messageListBox.Items.Add(messageString);

                //Populate steps list box.  They are added unchecked
                stepsCheckedListBox.Items.Clear();
                string stepsLine = File.ReadLines(taskTextLocation).Skip(4).Take(1).First();
                if (stepsLine != "")
                {
                    string[] stepsArr = stepsLine.Split('|');
                    foreach (string stepString in stepsArr)
                        if (stepString != "")
                            stepsCheckedListBox.Items.Add(stepString);
                }

                //Check steps. Only steps previously checked by the user are checked
                string checkedLine = File.ReadLines(taskTextLocation).Skip(5).Take(1).First();
                if (checkedLine != "")
                {
                    string[] checkedArr = checkedLine.Split('|');
                    for (int x = 0; x < checkedArr.Length; ++x)
                        if (checkedArr[x] != "")
                            stepsCheckedListBox.SetItemChecked(int.Parse(checkedArr[x]), true);
                }

                //Task Description
                descriptionRichTextBox.Text = "";
                string[] descriptions = File.ReadAllLines(taskTextLocation).Skip(6).ToArray();
                if (descriptions[0] != "")
                    foreach (string description in descriptions)
                        descriptionRichTextBox.Text += description + "\n";

                //Changes color of task status buttons based on task status
                normalButton.BackColor = SystemColors.Control;
                stuckButton.BackColor = SystemColors.Control;
                urgentButton.BackColor = SystemColors.Control;
                completeButton.BackColor = SystemColors.Control;
                //Set task color on buttons
                string taskNameString = File.ReadAllLines($@"{fileLocation}\Projects\{projectListBox.SelectedItem}\{projectListBox.SelectedItem}.txt").Skip(taskListBox.SelectedIndex).Take(1).First();
                string[] taskNameArr = taskNameString.Split('|');
                StatusButtonColor(int.Parse(taskNameArr[1]));
            }
        }

        //Task Status Buttons=================================================================================================
        //The next four methods use ButtonClicker() to change the status button color and task status
        private void NormalButton_Click(object sender, EventArgs e)
        {
            ButtonClicker(0);
        }

        private void StuckButton_Click(object sender, EventArgs e)
        {
            ButtonClicker(1);
        }

        private void UrgentButton_Click(object sender, EventArgs e)
        {
            ButtonClicker(2);
        }

        private void CompleteButton_Click(object sender, EventArgs e)
        {
            ButtonClicker(3);
        }

        private void ButtonClicker(int status)
        {
            //Updates button color and task status using StatusButtonColor() and ProjectLineChanger()
            //Serves as a bridge between the buttons and the necessary methods
            if (taskListBox.SelectedIndex > -1)
            {
                ProjectLineChanger(status, taskListBox.SelectedIndex + 1);
                StatusButtonColor(status);
            }
            else
                MessageBox.Show("Select a task first", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
        }

        public void StatusButtonColor(int color)
        {
            //Sets the color of the status buttons
            normalButton.BackColor = SystemColors.Control;
            stuckButton.BackColor = SystemColors.Control;
            urgentButton.BackColor = SystemColors.Control;
            completeButton.BackColor = SystemColors.Control;
            switch (color)
            {
                case 0:
                    break;
                case 1:
                    stuckButton.BackColor = Color.FromArgb(255, 50, 45);
                    break;
                case 2:
                    urgentButton.BackColor = Color.FromArgb(255, 255, 60);
                    break;
                case 3:
                    completeButton.BackColor = Color.FromArgb(90, 255, 60);
                    break;
                default:
                    break;
            }
        }

        public void ProjectLineChanger(int taskStatus, int lineToEdit)
        {
            //Updates the task status in the project file
            string fileName = $@"{fileLocation}\Projects\{projectListBox.SelectedItem.ToString()}\{projectListBox.SelectedItem.ToString()}.txt";
            string[] allLines = File.ReadAllLines(fileName);
            TaskNameFinder();
            allLines[lineToEdit - 1] = $"{taskName}" + "|" + taskStatus;
            File.WriteAllLines(fileName, allLines);
            //Update tasks list box names to reflect the task status
            switch (taskStatus)
            {
                case 0:
                    taskListBox.Items[taskListBox.SelectedIndex] = taskName;
                    break;
                case 1:
                    taskListBox.Items[taskListBox.SelectedIndex] = taskName + "  || STUCK";
                    break;
                case 2:
                    taskListBox.Items[taskListBox.SelectedIndex] = taskName + "  || URGENT";
                    break;
                case 3:
                    taskListBox.Items[taskListBox.SelectedIndex] = taskName + "  || COMPLETE";
                    break;
                default:
                    taskListBox.Items[taskListBox.SelectedIndex] = taskName;
                    break;
            }
        }



        //===========================================================================================================
        private void StepsCheckedListBox_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            //Save which steps are checked
            //Force the ItemCheck() method code to execute after the item has been checked, rather than before. 
            //By default, the control fires this event and then updates the list of checked items.  By using BeginInvoke(),
            //the control is forced to first update which items are checked before firing this event.
            this.BeginInvoke(new Action(() =>
            {
                string checkedIndices = "";
                foreach (int checkedIndexInt in stepsCheckedListBox.CheckedIndices)
                    checkedIndices += checkedIndexInt.ToString() + "|";
                LineChanger(checkedIndices, 6);
            }));
        }
