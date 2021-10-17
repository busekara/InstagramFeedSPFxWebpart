import * as React from "react";
import styles from "./InstagramFeed.module.scss";
import "./styles.css";

import { IInstagramFeedProps } from "./IInstagramFeedProps";
import {
  Shimmer,
  ShimmerElementsGroup,
  ShimmerElementType,
} from "office-ui-fabric-react/lib/Shimmer";
import { Fabric } from "office-ui-fabric-react/lib/Fabric";
import { mergeStyles } from "office-ui-fabric-react/lib/Styling";
import { MessageBar, MessageBarType } from "office-ui-fabric-react";

import { FocusZone } from "office-ui-fabric-react/lib/FocusZone";
import {
  IPersonaSharedProps,
  Persona,
} from "office-ui-fabric-react/lib/Persona";

import { List } from "office-ui-fabric-react/lib/List";

import * as strings from "InstagramFeedWebPartStrings";
import { IInstagramFeed } from "../models/IInstagramFeed";
import { IError } from "../models/IError";
import * as $ from "jquery";
import { uniqueId } from "lodash";

//import swiper and required css
import { Swiper, SwiperSlide } from "swiper/react";
import SwiperCore, { Navigation, Pagination } from "swiper";

import "swiper/swiper-bundle.css";
import "swiper/components/navigation/navigation.min.css";
import "swiper/components/pagination/pagination.min.css";

// install Swiper modules
SwiperCore.use([Pagination, Navigation]);

const shimmerWrapperClass = mergeStyles({
  padding: 2,
  selectors: {
    "& > .ms-Shimmer-container": {
      margin: "10px 0",
    },
  },
});

export default class InstagramFeed extends React.Component<
  IInstagramFeedProps,
  {
    error: IError;
    isLoaded: Boolean;
    items: IInstagramFeed;
    posts: any;
  }
> {
  constructor(props) {
    super(props);
    this.state = {
      error: null,
      isLoaded: false,
      items: null,
      posts: null,
    };
  }
  public componentDidMount() {
    this._loadData();
  }
  private async _loadData() {
    try {
      this.getData(0);
    } catch (exception) {
      console.warn(`${exception.code}-${exception.name}: ${exception.message}`);
      let failure: IError = {
        heading: strings.ExceptionHeading,
        message: strings.ExceptionMessage,
        status: 500,
      };
      // show exception
      this.setState({
        items: null,
        isLoaded: true,
        error: failure,
      });
    }
  }
  private getData(count) {
    if (this.props.userToken != null) {
      var params = {
        url: `https://graph.instagram.com/me/media?fields=media_type,media_url,caption,permalink&access_token=${this.props.userToken}`,
        container: "none",
      };

      try {
        let xhr = new XMLHttpRequest();
        xhr.open("GET", params.url);
        xhr.onload = () => {
          if (xhr.status === 200) {
            this.handleSuccess(xhr.responseText);
          } else if (xhr.status === 404) {
            this.handleFailure(xhr.status);
          } else {
            this.getData(1);
          }
        };
        xhr.send();
      } catch (exception) {
        if (0 === count) {
          this.getData(1);
        } else {
          throw exception;
        }
      }
    }
  }
  private handleSuccess(success) {
    let json = JSON.parse(success);
    if (json.data != null) {
      let data = json.data;
      this.setState({
        isLoaded: true,
        items: data,
        error: null,
      });
    }
  }
  private handleFailure(error) {
    let failure: IError = {
      heading: strings.ErrorHeading,
      message: strings.ErrorMessage,
      status: error,
    };
    // show error
    this.setState({
      items: null,
      isLoaded: true,
      error: failure,
    });
  }

  private _profilePersona = (): JSX.Element => {
    const examplePersona: IPersonaSharedProps = {
      imageUrl: require("../images/instagramIcon.png"),
      imageInitials: "IG",
      text: this.props.userFullName,
      secondaryText: this.props.accountName,
    };

    return <Persona {...examplePersona} />;
  };

  private _loadingShimmer = (): JSX.Element => {
    return (
      <div style={{ display: "flex" }}>
        <ShimmerElementsGroup
          shimmerElements={[
            { type: ShimmerElementType.line, width: 100, height: 100 },
            { type: ShimmerElementType.gap, width: 10, height: 100 },
            { type: ShimmerElementType.line, width: 100, height: 100 },
            { type: ShimmerElementType.gap, width: 10, height: 100 },
          ]}
        />
      </div>
    );
  };

  private _errorNotification = (): JSX.Element => {
    console.error(`${this.state.error.status}: ${this.state.error.message}`);
    return (
      <MessageBar
        messageBarType={MessageBarType.error}
        isMultiline={false}
        truncated={true}
      >
        <strong>{this.state.error.heading}</strong>
        {this.state.error.message ? " - " + this.state.error.message : ""}
      </MessageBar>
    );
  };

  private _onRenderCell = (item: any): JSX.Element => {
    const feeds = [];
    for (const post of item) {
      if (post.media_type == "VIDEO") {
        feeds.push(
          <SwiperSlide key={uniqueId("prefix-")}>
            <div data-is-focusable={true} role="img">
              <div className="content">
                <a target="_blank" href={post.permalink}>
                  <video src={post.media_url} className="postVideo" />
                  {post.caption ? <div>{post.caption}</div> : ""}
                </a>
              </div>
            </div>
          </SwiperSlide>
        );
      } else {
        feeds.push(
          <SwiperSlide key={uniqueId("prefix-")}>
            <div data-is-focusable={true} role="img">
              <div className="content">
                <a target="_blank" href={post.permalink}>
                  <img src={post.media_url} className="postImage" />
                  {post.caption ? <div>{post.caption}</div> : ""}
                </a>
              </div>
            </div>
          </SwiperSlide>
        );
      }
    }

    return (
      <Swiper
        spaceBetween={20}
        slidesPerView={3}
        navigation={true}
        pagination={{ clickable: true }}
        scrollbar={{ draggable: true }}
      >
        {feeds}
      </Swiper>
    );
  };
  private _onRenderCellFull = (item: any): JSX.Element => {
    const feeds = [];
    for (const post of item) {
      if (post.media_type == "VIDEO") {
        feeds.push(
          <SwiperSlide key={uniqueId("prefix-")}>
            <div data-is-focusable={true} role="img">
              <div className="content">
                <a target="_blank" href={post.permalink}>
                  <video src={post.media_url} className="postVideoFull" />
                  {post.caption ? <div>{post.caption}</div> : ""}
                </a>
              </div>
            </div>
          </SwiperSlide>
        );
      } else {
        feeds.push(
          <SwiperSlide key={uniqueId("prefix-")}>
            <div data-is-focusable={true} role="img">
              <div className="content">
                <a target="_blank" href={post.permalink}>
                  <img src={post.media_url} className="postImageFull" />
                  {post.caption ? <div>{post.caption}</div> : ""}
                </a>
              </div>
            </div>
          </SwiperSlide>
        );
      }
    }

    return (
      <Swiper
        spaceBetween={20}
        slidesPerView={3}
        navigation={true}
        pagination={{ clickable: true }}
        scrollbar={{ draggable: true }}
      >
        {feeds}
      </Swiper>
    );
  };
  public render(): JSX.Element {
    if (this.state.error) {
      return this._errorNotification();
    } else if (!this.state.isLoaded) {
      return (
        <Fabric className={shimmerWrapperClass}>
          <Shimmer customElementsGroup={this._loadingShimmer()} width={"45%"} />
        </Fabric>
      );
    } else if (this.state.items != null) {
      return (
        <div className="instagramFeed">
          {this.props.showIcon ? this._profilePersona() : ""}

          <FocusZone>
            <List
              items={[this.state.items]}
              onRenderCell={
                this.props.layoutOneThirdRight
                  ? this._onRenderCell
                  : this._onRenderCellFull
              }
            />
          </FocusZone>
        </div>
      );
    }
  }
}
