import * as React from 'react';
import type { ITeamsLeaveCalendarProps } from './ITeamsLeaveCalendarProps';
import { SPComponentLoader } from "@microsoft/sp-loader";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/webs";
import "@pnp/sp/site-users/web";
import "@pnp/sp/attachments";
import "@pnp/sp/presets/all";
import { Web } from "@pnp/sp/webs";
import { Calendar, momentLocalizer } from 'react-big-calendar';
import 'react-big-calendar/lib/css/react-big-calendar.css';
import * as moment from "moment";
import * as $ from "jquery";
import "../css/style.css";
import "../css/responsive.css"



SPComponentLoader.loadCss(`https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css`);
SPComponentLoader.loadCss(`https://cdnjs.cloudflare.com/ajax/libs/font-awesome/4.7.0/css/font-awesome.min.css`);
// SPComponentLoader.loadCss('https://remodigital.sharepoint.com/:f:/r/sites/Remo/RemoSolutions/VTM/SiteAssets/Visitor%20and%20Trainee%20Assets/css/style.css?v=2.9');
// SPComponentLoader.loadCss('https://remodigital.sharepoint.com/:f:/r/sites/Remo/RemoSolutions/VTM/SiteAssets/Visitor%20and%20Trainee%20Assets/css/responsivestyle.css?v=2.9');

var NewWeb: any;
const localizer = momentLocalizer(moment)


export interface FormState {
  UpcomingEvents: {
    id: string;
    title: string;
    start: any;
    end: any;
    Type: string;
    className?: string; // Add className for styling
  }[];
  SelectedEventItems: any[];
  SelectedPermissionItem: any[];
  CurrentView: any;
}

export default class TeamsLeaveCalendar extends React.Component<ITeamsLeaveCalendarProps, FormState, {}> {
  public constructor(props: ITeamsLeaveCalendarProps, state: FormState) {
    super(props);
    this.state = {
      // UpcomingEvents: [{
      //   id: "1",
      //   title: "Test",
      //   start: "08/07/2024",
      //   end: "08/10/2024",
      // }],
      UpcomingEvents: [],
      SelectedEventItems: [],
      CurrentView: "month",
      SelectedPermissionItem: []
    }
    NewWeb = Web("" + this.props.siteurl + "")
    this.handleNavigate = this.handleNavigate.bind(this);

  }
  public componentDidMount() {
    this.getLeaveRequestDetails()
  }
  public toggleLogout() {
    $(".btn-log-out").toggle();
  }
  public getEventStyle(event: any) {
    if (event.Type == "Leave") {
      return {
        className: 'leave_request', // Add a class for styling
      };
    }
    else if (event.Type == "Permission") {
      return {
        className: 'permission_request', // Add a class for styling
      };
    }
    return {};
  };
  public async handleEventClick(event: any, e: React.SyntheticEvent): Promise<void> {
    e.preventDefault();
    console.log("Event:", event)
    var ID = event.id
    var Type = event.Type

    if (Type == "Leave") {
      const items = await NewWeb.lists.getByTitle("LeaveRequest")
        .items.select("*").filter(`ID eq ${ID}`).get();
      this.setState({
        SelectedEventItems: items
      });
      $("#LeaveRequest-table-details").show();
    } else if (Type == "Permission") {
      const items = await NewWeb.lists.getByTitle("EmployeePermission")
        .items.select("*").filter(`ID eq ${ID}`).get();
      this.setState({
        SelectedPermissionItem: items
      });
      $("#PermissionRequest-table-details").show();
    }
  }
  public getLeaveRequestDetails() {
    NewWeb.lists.getByTitle("LeaveRequest").items.select("*").filter(`Status eq 'Approved'`).getAll()
      .then((items: any) => {
        console.log(items)
        if (items.length !== 0) {
          const formattedEvents = items.map((item: any) => {
            // Parse the StartDate and EndDate from the backend list
            const startDate = moment(item.StartDate).format("MM/DD/YYYY");
            const endDate = moment(item.EndDate).add(1, 'day').format("MM/DD/YYYY");
            return {
              id: item.ID,
              title: `${item.LeaveType} - ${item.Requester}`,  // Ensure title is set correctly
              start: startDate,
              end: endDate,
              Type: "Leave"
            };
          });

          this.setState({
            UpcomingEvents: formattedEvents
          });
          this.getPermissionRequestDetails(formattedEvents)
        }

      });
  }
  public getPermissionRequestDetails(existingEvents: any[]) {
    NewWeb.lists.getByTitle("EmployeePermission").items.select("*").filter(`Status eq 'Approved'`).getAll()
      .then((items: any) => {
        if (items.length !== 0) {
          const formattedEvents = items.map((item: any) => {
            // Parse the StartDate and EndDate from the backend list
            const startDate = moment(item.timefromwhen).format("DD/MM/YYYY");
            const endDate = moment(item.TimeUpto).format("DD/MM/YYYY");
            return {
              id: item.ID,
              title: `Permission - ${item.Requester}`,  // Ensure title is set correctly
              start: startDate,
              end: endDate,
              Type: "Permission"
            };
          });
          // Merge existing events with new permission events
          const allEvents = existingEvents.concat(formattedEvents);
          this.setState({
            UpcomingEvents: allEvents
          });
          console.log("Upcoming Events:", this.state.UpcomingEvents);

        }

      });
  }
  public handleNavigate(newDate: any) {
    var handler = this;
    handler.setState({
      CurrentView: "month",
    })
    $("#LeaveRequest-table-details").hide()
    $("#PermissionRequest-table-details").hide();
  }
  public closeTable(Type: any) {
    if (Type == "Leave") {
      $("#LeaveRequest-table-details").hide();
    } else if (Type == "Permission") {
      $("#PermissionRequest-table-details").hide();
    }
  }
  public render(): React.ReactElement<ITeamsLeaveCalendarProps> {
    // const {
    //   description,
    //   isDarkTheme,
    //   environmentMessage,
    //   hasTeamsContext,
    //   userDisplayName
    // } = this.props;

    return (
      <>
        <section id='load-content'>
          <div className="store-section user-calendar">
            <div className="row store-wrap user_calendar">
              <Calendar
                localizer={localizer}
                events={this.state.UpcomingEvents}
                startAccessor="start"
                endAccessor="end"
                // views=""
                view={this.state.CurrentView}
                onView={(view) => this.setState({ CurrentView: view })}
                date=""
                eventPropGetter={this.getEventStyle}
                style={{ height: 405 }}
                onNavigate={this.handleNavigate}
                tooltipAccessor="Type"
                onSelectEvent={(event, e) => this.handleEventClick(event, e)}
              />
            </div>
          </div>
          <div className='table-popup' style={{ display: "none" }} id='LeaveRequest-table-details'>
            <div className='table-overlay_popup'>
              <div className="manual-booking-table view-event-table user-calendar">
                <div className="table-responsive" id="table-content">
                  <h4 className="events_title">Event Details</h4>
                  <div className="popup_cancel" onClick={() => this.closeTable("Leave")}>
                    <img src={require("../Images/close-icon.svg")} />
                  </div>
                  <table className="table" id="table-example">
                    <thead>
                      <tr>
                        <th>S.No</th>
                        <th>Requestor Name</th>
                        <th>Status</th>
                        <th>Leave Type</th>
                        <th>Half Day / Full Day</th>
                        <th>Start Date</th>
                        <th>End Date</th>
                        <th>Reason</th>
                        <th>Manager Comments</th>
                      </tr>
                    </thead>
                    <tbody>
                      {this.state.SelectedEventItems && this.state.SelectedEventItems.map((item, key) => {
                        return (
                          <tr>
                            <td>{key + 1}</td>
                            <td>{item.Requester}</td>
                            <td>{item.Status}</td>
                            <td>{item.LeaveType}</td>
                            <td>{item.Day}</td>
                            <td>{item.StartDate}</td>
                            <td>{item.EndDate}</td>
                            <td>{item.Reason}</td>
                            <td>{item.ManagerComments}</td>

                          </tr>
                        )
                      })}

                    </tbody>

                  </table>

                </div>
              </div>
            </div>
          </div>
          <div className='table-popup' style={{ display: "none" }} id='PermissionRequest-table-details'>
            <div className='table-overlay_popup'>
              <div className="manual-booking-table view-event-table user-calendar">
                <div className="table-responsive" id="table-content">
                  <h4 className="events_title">Event Details</h4>
                  <div className="popup_cancel" onClick={() => this.closeTable("Permission")}>
                    <img src={require("../Images/close-icon.svg")} />
                  </div>
                  <table className="table" id="table-example">
                    <thead>
                      <tr>
                        <th>S.No</th>
                        <th>Requestor Name</th>
                        <th>Status</th>
                        <th>Start Time</th>
                        <th>End Time</th>
                        <th>Permission Hours</th>
                        <th>Reason</th>
                        <th>Manager Comments</th>
                      </tr>
                    </thead>
                    <tbody>
                      {this.state.SelectedPermissionItem && this.state.SelectedPermissionItem.map((item, key) => {
                        return (
                          <tr>
                            <td>{key + 1}</td>
                            <td>{item.Requester}</td>
                            <td>{item.Status}</td>
                            <td>{item.timefromwhen}</td>
                            <td>{item.TimeUpto}</td>
                            <td>{item.PermissionHour}</td>
                            <td>{item.Reason}</td>
                            <td>{item.ManagerComments}</td>

                          </tr>
                        )
                      })}

                    </tbody>

                  </table>

                </div>
              </div>
            </div>
          </div>
        </section>
      </>
    );
  }
}
