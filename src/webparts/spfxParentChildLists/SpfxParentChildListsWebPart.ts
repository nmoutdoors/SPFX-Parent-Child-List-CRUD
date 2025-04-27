import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import * as strings from 'SpfxParentChildListsWebPartStrings';
import styles from './components/SpfxParentChildLists.module.scss';

// Add enum for Status options
enum ProjectStatus {
  NotStarted = 'Not Started',
  InProgress = 'In Progress',
  Completed = 'Completed'
}

interface IMajorProject {
  ID: number;
  Title: string;
  Description: string;
  Status: ProjectStatus;
}

interface IMajorProjectEvent {
  ID: number;
  Title: string;
  Start: string;  // Changed from Date to string
  End: string;    // Changed from Date to string
  Status: ProjectStatus;
  MajorProject: {
    Id: number;
    Title: string;
  };
}

interface IEditDialogState {
  isOpen: boolean;
  project: IMajorProject | undefined;
}

interface IStatusOption {
  key: ProjectStatus;  // Changed from string to ProjectStatus
  text: string;
}

// Add this interface to define the shape of raw SharePoint event data
interface ISharePointEventData {
  ID: number;
  Title: string;
  Start: string;
  End: string;
  Status: string;
  MajorProject: {
    Id: number;
    Title: string;
  };
}

// Add this interface for event editing state
interface IEventEditState {
  isEditing: boolean;
  eventId: number | undefined;  // Changed from null to undefined
}

export interface ISpfxParentChildListsWebPartProps {
  description: string;
}

export default class SpfxParentChildListsWebPart extends BaseClientSideWebPart<ISpfxParentChildListsWebPartProps> {
  private projects: IMajorProject[] = [];
  private projectEvents: IMajorProjectEvent[] = [];
  private dialogState: IEditDialogState = {
    isOpen: false,
    project: undefined
  };

  private editedProject: IMajorProject | undefined = undefined;

  private statusOptions: IStatusOption[] = [
    { key: ProjectStatus.NotStarted, text: 'Not Started' },
    { key: ProjectStatus.InProgress, text: 'In Progress' },
    { key: ProjectStatus.Completed, text: 'Completed' }
  ];

  // Add this property to track event editing state
  private eventEditState: IEventEditState = {
    isEditing: false,
    eventId: undefined  // Changed from null to undefined
  };

  private async getMajorProjects(): Promise<IMajorProject[]> {
    const response: SPHttpClientResponse = await this.context.spHttpClient.get(
      `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('MajorProjects')/items?$select=ID,Title,Description,Status`,
      SPHttpClient.configurations.v1
    );

    const data = await response.json();
    return data.value;
  }

  private async updateMajorProject(project: IMajorProject): Promise<void> {
    const headers = {
      'X-HTTP-Method': 'MERGE',
      'IF-MATCH': '*'
    };

    await this.context.spHttpClient.post(
      `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('MajorProjects')/items(${project.ID})`,
      SPHttpClient.configurations.v1,
      {
        headers: headers,
        body: JSON.stringify({
          Title: project.Title,
          Description: project.Description,
          Status: project.Status
        })
      }
    );
  }

  // Add method to update event
  private async updateMajorProjectEvent(event: IMajorProjectEvent): Promise<void> {
    const headers = {
      'X-HTTP-Method': 'MERGE',
      'IF-MATCH': '*'
    };

    await this.context.spHttpClient.post(
      `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('MajorProjectEvents')/items(${event.ID})`,
      SPHttpClient.configurations.v1,
      {
        headers: headers,
        body: JSON.stringify({
          Title: event.Title,
          Start: event.Start,
          End: event.End,
          Status: event.Status
        })
      }
    );
  }

  private getCardHTML(project: IMajorProject): string {
    const truncatedDescription = project.Description 
      ? project.Description.substring(0, 100) + (project.Description.length > 100 ? '...' : '')
      : 'No description available';

    return `
      <div class="project-card" style="
        border: 1px solid #ccc;
        border-radius: 4px;
        margin: 10px;
        padding: 0;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        background: white;
        width: 300px;
        display: inline-block;
        vertical-align: top;
        cursor: pointer;
      " data-project-id="${project.ID}">
        <div class="card-header" style="
          background: #0078d4;
          color: white;
          padding: 10px;
          font-weight: bold;
          border-radius: 4px 4px 0 0;
        ">
          ${project.Title}
        </div>
        <div class="card-body" style="
          padding: 15px;
          min-height: 80px;
        ">
          ${truncatedDescription}
        </div>
        <div class="card-footer" style="
          padding: 10px;
          background: #f8f8f8;
          border-top: 1px solid #eee;
          border-radius: 0 0 4px 4px;
        ">
          Status: ${project.Status || 'Not set'}
        </div>
      </div>
    `;
  }

  private getStatusStyle(status: ProjectStatus): { backgroundColor: string; color: string } {
    switch (status) {
      case ProjectStatus.NotStarted:
        return {
          backgroundColor: '#FFCCCC',  // Darker red background
          color: '#B71C1C'            // Even darker red for text
        };
      case ProjectStatus.InProgress:
        return {
          backgroundColor: '#FFF176',  // More yellow background
          color: '#664200'            // Darker brown for text
        };
      case ProjectStatus.Completed:
        return {
          backgroundColor: '#C8E6C9',  // Darker green background
          color: '#1B5E20'            // Even darker green for text
        };
      default:
        return {
          backgroundColor: '#E1DFDD',  // Light grey
          color: '#323130'            // Dark grey for text
        };
    }
  }

  private renderEditDialog(): string {
    if (!this.dialogState.isOpen || !this.dialogState.project) return '';

    const project = this.dialogState.project;
    const events = this.projectEvents || [];
    
    // Use the events variable to render project events
    const eventsHtml = events.map(event => {
        const isEditing = this.eventEditState.isEditing && this.eventEditState.eventId === event.ID;
        const statusStyle = this.getStatusStyle(event.Status);
        
        if (isEditing) {
            return `
                <div class="event-item" data-event-id="${event.ID}" style="
                    margin: 10px 0;
                    padding: 10px;
                    border: 1px solid #ccc;
                    background: ${statusStyle.backgroundColor};
                    box-shadow: 0 3.2px 7.2px 0 rgba(0, 0, 0, 0.132), 0 0.6px 1.8px 0 rgba(0, 0, 0, 0.108);
                    border-radius: 4px;
                ">
                    <input type="text" class="event-title" value="${event.Title}" style="width: 100%; margin-bottom: 5px;">
                    <input type="date" class="event-start" value="${event.Start.split('T')[0]}" style="margin-right: 10px;">
                    <input type="date" class="event-end" value="${event.End.split('T')[0]}" style="margin-right: 10px;">
                    <select class="event-status">
                        ${this.statusOptions.map(option => `
                            <option value="${option.key}" ${event.Status === option.key ? 'selected' : ''}>
                                ${option.text}
                            </option>
                        `).join('')}
                    </select>
                    <button class="save-event-edit" data-event-id="${event.ID}">Save</button>
                    <button class="cancel-event-edit">Cancel</button>
                </div>
            `;
        }

        return `
            <div class="event-item" style="
                margin: 10px 0;
                padding: 10px;
                border: 1px solid #ccc;
                background: ${statusStyle.backgroundColor};
                box-shadow: 0 3.2px 7.2px 0 rgba(0, 0, 0, 0.132), 0 0.6px 1.8px 0 rgba(0, 0, 0, 0.108);
                border-radius: 4px;
            ">
                <strong>${event.Title}</strong><br>
                ${this.formatDate(event.Start)} - ${this.formatDate(event.End)}<br>
                Status: ${event.Status}
                <button class="edit-event-button" data-event-id="${event.ID}" style="float: right;">Edit</button>
            </div>
        `;
    }).join('');

    const statusOptionsHtml = this.statusOptions
        .map(option => `
            <option value="${option.key}" ${project.Status === option.key ? 'selected' : ''}>
                ${option.text}
            </option>
        `)
        .join('');

    return `
        <div class="dialog-overlay" style="
          position: fixed;
          top: 0;
          left: 0;
          right: 0;
          bottom: 0;
          background-color: rgba(0, 0, 0, 0.5);
          display: flex;
          justify-content: center;
          align-items: center;
          z-index: 1000;
        ">
          <div class="dialog-content" style="
            background: white;
            padding: 20px;
            border-radius: 4px;
            box-shadow: 0 4px 8px rgba(0, 0, 0, 0.2);
            width: 80%;
            max-width: 800px;
            max-height: 90vh;
            overflow-y: auto;
            position: relative;
          ">
            <h2>Edit Project</h2>
            
            <div style="margin-bottom: 15px;">
                <label>Title:</label>
                <input type="text" class="project-title" value="${project.Title}" style="width: 100%; margin-top: 5px;">
            </div>
            
            <div style="margin-bottom: 15px;">
                <label>Description:</label>
                <textarea class="project-description" style="width: 100%; height: 100px; margin-top: 5px;">${project.Description || ''}</textarea>
            </div>
            
            <div style="margin-bottom: 15px;">
                <label>Status:</label>
                <select class="project-status" style="width: 100%; margin-top: 5px;">
                    ${statusOptionsHtml}
                </select>
            </div>

            <div style="margin-top: 20px;">
                <h3>Related Events Stored in the MajorProjectEvents list</h3>
                ${eventsHtml}
            </div>

            <div style="text-align: right; margin-top: 20px;">
                <button class="cancel-button" style="margin-right: 10px;">Cancel</button>
                <button class="save-button">Save</button>
            </div>
          </div>
        </div>
    `;
  }

  private switchTab(tabName: 'details' | 'events'): void {
    // Get all tab buttons and content
    const tabButtons = this.domElement.getElementsByClassName('tab-button');
    const tabContents = this.domElement.getElementsByClassName('tab-content');

    // Hide all content and remove active class from buttons
    Array.from(tabContents).forEach((content: Element) => {
      (content as HTMLElement).style.display = 'none';
    });

    Array.from(tabButtons).forEach((button: Element) => {
      button.classList.remove('active');
      (button as HTMLElement).style.color = '#605e5c';
      (button as HTMLElement).style.fontWeight = 'normal';
    });

    // Show selected content and activate button
    const selectedContent = this.domElement.querySelector(`#${tabName}Content`);
    const selectedTab = this.domElement.querySelector(`#${tabName}Tab`);

    if (selectedContent) {
      (selectedContent as HTMLElement).style.display = 'block';
    }

    if (selectedTab) {
      selectedTab.classList.add('active');
      (selectedTab as HTMLElement).style.color = '#0078d4';
      (selectedTab as HTMLElement).style.fontWeight = '600';
    }
  }

  private attachEventHandlers(): void {
    // Project card click handlers
    const projectCards = this.domElement.getElementsByClassName('project-card');
    Array.from(projectCards).forEach((card: Element) => {
      card.addEventListener('click', async (e: Event) => {
        const projectId = parseInt((e.currentTarget as HTMLElement).getAttribute('data-project-id') || '0');
        try {
          await this.openEditDialog(projectId);
        } catch (error: unknown) {
          console.error('Error opening dialog:', error);
        }
      });
    });

    // Dialog save button handler
    const saveButton = this.domElement.querySelector('.save-button');
    if (saveButton) {
      saveButton.addEventListener('click', async () => {
        await this.saveProject().catch(error => {
          console.error('Error saving project:', error);
        });
      });
    }

    // Dialog cancel button handler
    const cancelButton = this.domElement.querySelector('.cancel-button');
    if (cancelButton) {
      cancelButton.addEventListener('click', async () => {
        await this.closeEditDialog().catch(error => {
          console.error('Error closing dialog:', error);
        });
      });
    }

    // Event save button handlers
    const saveEventButtons = this.domElement.getElementsByClassName('save-event-edit');
    Array.from(saveEventButtons).forEach((button: Element) => {
      button.addEventListener('click', async (e: Event) => {
        e.preventDefault();
        e.stopPropagation();
        const eventId = parseInt((e.currentTarget as HTMLElement).getAttribute('data-event-id') || '0');
        await this.saveEventEdit(eventId).catch(error => {
          console.error('Error saving event:', error);
        });
      });
    });

    // Event cancel button handlers
    const cancelEventButtons = this.domElement.getElementsByClassName('cancel-event-edit');
    Array.from(cancelEventButtons).forEach((button: Element) => {
      button.addEventListener('click', (e: Event) => {
        e.preventDefault();
        e.stopPropagation();
        this.cancelEventEdit().catch(error => {
          console.error('Error canceling event edit:', error);
        });
      });
    });

    // Event edit button handlers
    const editEventButtons = this.domElement.getElementsByClassName('edit-event-button');
    Array.from(editEventButtons).forEach((button: Element) => {
      button.addEventListener('click', (e: Event) => {
        e.preventDefault();
        e.stopPropagation();
        const eventId = parseInt((e.currentTarget as HTMLElement).getAttribute('data-event-id') || '0');
        this.startEventEdit(eventId).catch(error => {
          console.error('Error starting event edit:', error);
        });
      });
    });

    // Tab button handlers
    const tabButtons = this.domElement.getElementsByClassName('tab-button');
    Array.from(tabButtons).forEach((button: Element) => {
      button.addEventListener('click', (e: Event) => {
        const tabName = (e.currentTarget as HTMLElement).id.replace('Tab', '') as 'details' | 'events';
        this.switchTab(tabName);
      });
    });

    // Add click handler for dialog overlay
    const dialogOverlay = this.domElement.querySelector('.dialog-overlay');
    const dialogContent = this.domElement.querySelector('.dialog-content');
    
    if (dialogOverlay && dialogContent) {
      dialogOverlay.addEventListener('click', async (e: Event) => {
        // Check if the click was on the overlay itself and not its children
        if (e.target === dialogOverlay) {
          await this.closeEditDialog().catch(error => {
            console.error('Error closing dialog:', error);
          });
        }
      });
    }
  }

  private async openEditDialog(projectId: number): Promise<void> {
    try {
      console.log('Opening dialog for project ID:', projectId);
      
      // Initialize projectEvents as empty array before fetching
      this.projectEvents = [];
      
      const [projects, events] = await Promise.all([
        this.getMajorProjects(),
        this.getMajorProjectEvents(projectId)
      ]);

      console.log('Received projects:', projects);
      console.log('Received events:', events);

      const matchingProjects = projects.filter((p: IMajorProject) => p.ID === projectId);
      const project = matchingProjects.length > 0 ? matchingProjects[0] : undefined;
      
      if (project) {
        this.dialogState = {
          isOpen: true,
          project: { ...project }
        };
        this.editedProject = { ...project };
        this.projectEvents = events;
        console.log('Set project events:', this.projectEvents);
        await this.render();
      } else {
        console.error('No matching project found for ID:', projectId);
      }
    } catch (error) {
      console.error('Error opening dialog:', error);
      this.projectEvents = [];
    }
  }

  private async closeEditDialog(): Promise<void> {
    this.dialogState = {
      isOpen: false,
      project: undefined
    };
    this.editedProject = undefined;
    await this.render();
  }

  private async saveProject(): Promise<void> {
    if (!this.editedProject || !this.dialogState.project) return;

    const titleInput = this.domElement.querySelector('.project-title') as HTMLInputElement;
    const descriptionInput = this.domElement.querySelector('.project-description') as HTMLTextAreaElement;
    const statusInput = this.domElement.querySelector('.project-status') as HTMLSelectElement;

    if (titleInput && descriptionInput && statusInput) {
      // Create a properly typed updated project
      const updatedProject: IMajorProject = {
        ID: this.editedProject.ID,
        Title: titleInput.value,
        Description: descriptionInput.value || '',  // Ensure non-null value
        Status: statusInput.value as ProjectStatus
      };

      try {
        await this.updateMajorProject(updatedProject);
        
        // Update the local projects array with the properly typed project
        this.projects = this.projects.map(project => 
          project.ID === updatedProject.ID ? updatedProject : project
        );

        await this.closeEditDialog();
        await this.render();
      } catch (error) {
        console.error('Error saving project:', error);
      }
    } else {
      console.error('Could not find one or more input elements:', {
        titleFound: !!titleInput,
        descriptionFound: !!descriptionInput,
        statusFound: !!statusInput
      });
    }
  }

  // Add these methods to handle event editing
  private async startEventEdit(eventId: number): Promise<void> {
    try {
      this.eventEditState = {
        isEditing: true,
        eventId: eventId
      };
      this.render();
    } catch (error: unknown) {
      console.error('Error starting event edit:', error);
    }
  }

  private async cancelEventEdit(): Promise<void> {
    this.eventEditState = {
      isEditing: false,
      eventId: undefined  // Changed from null to undefined
    };
    await this.render();
  }

  private async saveEventEdit(eventId: number): Promise<void> {
    console.log('saveEventEdit called with ID:', eventId); // Debug log
    const eventCard = this.domElement.querySelector(`.event-item[data-event-id="${eventId}"]`);
    if (!eventCard) {
      console.log('Event card not found'); // Debug log
      return;
    }

    const titleInput = eventCard.querySelector('.event-title') as HTMLInputElement;
    const startInput = eventCard.querySelector('.event-start') as HTMLInputElement;
    const endInput = eventCard.querySelector('.event-end') as HTMLInputElement;
    const statusInput = eventCard.querySelector('.event-status') as HTMLSelectElement;

    if (!titleInput || !startInput || !endInput || !statusInput) {
      console.log('One or more inputs not found'); // Debug log
      return;
    }

    const event = this.projectEvents.find(e => e.ID === eventId);
    if (!event) {
      console.log('Event not found in projectEvents'); // Debug log
      return;
    }

    const updatedEvent: IMajorProjectEvent = {
      ...event,
      Title: titleInput.value,
      Start: startInput.value,
      End: endInput.value,
      Status: statusInput.value as ProjectStatus
    };

    try {
      console.log('Updating event:', updatedEvent); // Debug log
      await this.updateMajorProjectEvent(updatedEvent);
      
      // Update local state
      this.projectEvents = this.projectEvents.map(e => 
        e.ID === eventId ? updatedEvent : e
      );
      
      this.eventEditState = {
        isEditing: false,
        eventId: undefined
      };
      
      await this.render();
    } catch (error) {
      console.error('Error updating event:', error);
    }
  }

  protected render(): void {
    if (!this.domElement) return;

    this.domElement.innerHTML = `
      <div class="${styles.spfxParentChildLists}">
        <div style="
          margin-bottom: 20px;
          padding: 15px;
          background-color: #f3f2f1;
          border-radius: 4px;
        ">
          <h1 style="
            margin: 0;
            color: #323130;
            font-size: 24px;
            font-weight: 600;
          ">Editing Parent and Child List Items in One Modal</h1>
        </div>

        <div class="ms-Grid">
          <div class="ms-Grid-row">
            <div class="ms-Grid-col ms-sm12">
              <div class="projects-container" style="
                display: flex;
                flex-wrap: wrap;
                gap: 20px;
                padding: 20px;
              ">
                ${this.projects.map(project => this.getCardHTML(project)).join('')}
              </div>
            </div>
          </div>
        </div>

        ${this.dialogState.isOpen ? this.renderEditDialog() : ''}
      </div>
    `;

    this.attachEventHandlers();
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }

  private async getMajorProjectEvents(projectId: number): Promise<IMajorProjectEvent[]> {
    try {
      const baseUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('MajorProjectEvents')/items`;
      
      const filteredUrl = `${baseUrl}?` +
        `$select=ID,Title,Start,End,Status,MajorProject/Id,MajorProject/Title` +
        `&$expand=MajorProject` +
        `&$filter=MajorProject/Id eq ${projectId}` +
        `&$orderby=Start asc`;

      const response = await this.context.spHttpClient.get(
        filteredUrl,
        SPHttpClient.configurations.v1
      );

      if (!response.ok) {
        console.error('Response not OK:', response.status, response.statusText);
        const errorText = await response.text();
        console.error('Error response:', errorText);
        return [];
      }

      const data = await response.json();

      if (data.value && Array.isArray(data.value)) {
        return data.value.map((event: ISharePointEventData) => ({
          ...event,
          Start: event.Start,
          End: event.End,
          Status: event.Status as ProjectStatus
        }));
      }

      return [];
    } catch (error) {
      console.error('Error in getMajorProjectEvents:', error);
      return [];
    }
  }

  private formatDate(dateString: string): string {
    return new Date(dateString).toLocaleDateString('en-US', {
      year: 'numeric',
      month: 'short',
      day: 'numeric'
    });
  }

  // Make sure you have this method to fetch projects when the web part loads
  protected onInit(): Promise<void> {
    return super.onInit().then(async () => {
      try {
        this.projects = await this.getMajorProjects();
        await this.render();
      } catch (error: unknown) {
        console.error('Error initializing web part:', error);
      }
    });
  }
}


